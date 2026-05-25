# --- Variables/Credentials ---
$tenantId      = Get-AutomationVariable -Name "TenantID"
$credentials   = Get-AutomationPSCredential -Name 'D365Offboarding'
$clientId      = $credentials.GetNetworkCredential().Username
$clientSecret  = $credentials.GetNetworkCredential().Password
$foBaseUrl     = "" #D365 environment url
$odataEntity   = "SystemUsers"

$workspaceId   = "" # Workspace id for sentinel
$WorkspaceSubscriptionId = "" # workspace subscription id for sentinel

# Connect Graph & Az
Connect-MgGraph -Identity
Connect-AzAccount -Identity
Set-AzContext -SubscriptionId $WorkspaceSubscriptionId  # add this line
# --- Get D365 Access Token ---
$tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $clientId
    scope         = "$foBaseUrl/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

try {
    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body
    $accessToken   = $tokenResponse.access_token
} catch {
    Write-Output "Failed to get access token: $_"
}

# --- Get D365 Users ---
$odataUrl = "$foBaseUrl/data/$odataEntity/"
$headers = @{
    "Authorization" = "Bearer $accessToken"
    "Accept"        = "application/json"
}

try {
    $users         = Invoke-RestMethod -Method Get -Uri $odataUrl -Headers $headers
    $filteredUsers = $users.value
    Write-Output "Retrieved $($filteredUsers.Count) records from $odataEntity"
} catch {
    Write-Output "OData request failed: $_"
}

# --- Query Sentinel via REST ---
Write-Output "Querying Sentinel for sign-in history..."

$query = @"
SigninLogs
| where AppDisplayName == 'Microsoft Dynamics ERP' or AppDisplayName == 'D365 PROD Tasklet Integration'
| summarize LastSignIn = max(TimeGenerated), AppLastUsed = arg_max(TimeGenerated, AppDisplayName) by UserPrincipalName
"@

# For PowerShell 7+ (convert SecureString to plain text)
$laToken = (Get-AzAccessToken -ResourceUrl "https://api.loganalytics.io").Token
$laTokenPlain = [System.Net.NetworkCredential]::new("", $laToken).Password

$queryResponse = Invoke-RestMethod `
    -Uri "https://api.loganalytics.io/v1/workspaces/$workspaceId/query" `
    -Method Post `
    -Headers @{
        Authorization  = "Bearer $laTokenPlain"
        "Content-Type" = "application/json"
    } `
    -Body (@{ query = $query } | ConvertTo-Json)

# Fetch all role assignments from D365
$roleUrl = "$foBaseUrl/data/SecurityUserRoles"
try {
    $roleResponse = Invoke-RestMethod -Method Get -Uri $roleUrl -Headers $headers
    $allRoles = $roleResponse.value
    Write-Output "Retrieved $($allRoles.Count) role assignments"
} catch {
    Write-Output "Failed to get role assignments: $_"
}

# Build a set of users who have the roles to skip
$skipRoles = @(
    "Warehouse mobile device user",
    "HS-Central-Advanced WMS Worker"
)

$usersToSkip = $allRoles `
    | Where-Object { $skipRoles -contains $_.SecurityRoleName } `
    | Select-Object -ExpandProperty UserId `
    | ForEach-Object { $_.ToLower() } `
    | Sort-Object -Unique

Write-Output "Users to skip due to WMS roles: $($usersToSkip.Count)"

# --- Parse Sentinel Response into Lookup Dictionary ---
$columns = $queryResponse.tables[0].columns | ForEach-Object { $_.name }
$rows    = $queryResponse.tables[0].rows

$dynamicsLookup = @{}
$taskletLookup  = @{}

foreach ($row in $rows) {
    $upn      = $row[$columns.IndexOf("UserPrincipalName")]
    $lastSign = $row[$columns.IndexOf("LastSignIn")]
    $appName  = $row[$columns.IndexOf("AppDisplayName")]

    if ($upn) {
        if ($appName -eq 'Microsoft Dynamics ERP') {
            $dynamicsLookup[$upn.ToLower()] = [datetime]$lastSign
        } elseif ($appName -eq 'D365 PROD Tasklet Integration') {
            $taskletLookup[$upn.ToLower()] = [datetime]$lastSign
        }
    }
}
# --- Process Users ---
$cutOff       = (Get-Date).AddDays(-60)
$InactiveUsers = @()

foreach ($user in $filteredUsers) {

    if ([string]::IsNullOrWhiteSpace($user.Email)) {
        Write-Output "Skipping $($user.UserID) - no email"
        continue
    }

    # Skip WMS role users
    if ($usersToSkip -contains $user.UserID.ToLower()) {
        Write-Output "Skipping $($user.UserID) - WMS role"
        continue
    }

    $email      = $user.Email.ToLower()
    $entraUser  = Get-MgUser -UserId $email -Property "createdDateTime" -ErrorAction SilentlyContinue
    $swedishDate = "N/A"
    $parsedDate  = $null

    if ($entraUser.CreatedDateTime) {
        $parsedDate  = [datetime]$entraUser.CreatedDateTime
        $swedishDate = $parsedDate.ToString("yyyy-MM-dd")
    }

    # Skip non-hydroscand emails, service accounts, or accounts created less than 60 days ago
    if (($email -notmatch "hydroscand") -or ($user.UserID -match "svc") -or ($parsedDate -ne $null -and $parsedDate -gt $cutOff)) {
        Write-Output "Skipping $($user.UserID)"
        continue
    }

    $lastSignInDynamics = $dynamicsLookup[$email]
    $lastSignInTasklet  = $taskletLookup[$email]

    # If active in Tasklet skip regardless of Dynamics activity
    if ($lastSignInTasklet -ne $null -and $lastSignInTasklet -ge $cutOff) {
        Write-Output "Skipping $($user.UserID) - active in Tasklet"
        continue
    }

    if (-not $entraUser) {
        $InactiveUsers += [pscustomobject]@{
            UserId             = $user.UserID
            Email              = $user.Email
            LastSignIn         = $null
            AppLastUsed        = "N/A"
            AccountCreatedDate = "N/A"
            "Enabled in D365"  = $user.Enabled
            Reason             = "User Not Found in Entra"
        }
    } elseif ($null -eq $lastSignInDynamics -or $lastSignInDynamics -lt $cutOff) {
        $InactiveUsers += [pscustomobject]@{
            UserId             = $user.UserID
            Email              = $user.Email
            LastSignIn         = if ($lastSignInDynamics) { $lastSignInDynamics.ToString("yyyy-MM-dd") } else { "No record in Sentinel" }
            AppLastUsed        = "Microsoft Dynamics ERP"
            AccountCreatedDate = $swedishDate
            "Enabled in D365"  = $user.Enabled
            Reason             = "Inactive > 60 Days in Dynamics"
        }
    }
}

Write-Output "Inactive users count: $($InactiveUsers.Count)"

# --- Build CSV ---
try {
    $CsvContent = $InactiveUsers | ConvertTo-Csv -NoTypeInformation
    $CsvBytes   = [System.Text.Encoding]::UTF8.GetBytes($CsvContent -join "`n")
    $CsvBase64  = [System.Convert]::ToBase64String($CsvBytes)
} catch {
    Write-Error "Failed to convert data to CSV: $_"
    exit 1
}

# --- Send Email ---
$emailPayload = @{
    message = @{
        subject = "Inactive Users found in PROD"
        body = @{
            contentType = "Text"
            content     = "Inactive Users in PROD"
        }
        toRecipients = @(
            @{ emailAddress = @{ address = "email@domain.com" } }
        )
        attachments = @(
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name          = "InactiveUsersPROD.csv"
                contentType   = "text/csv"
                contentBytes  = $CsvBase64
            }
        )
    }
    saveToSentItems = "true"
}

Send-MgUserMail -UserId "email@domain.com" -BodyParameter $emailPayload
# --- Variables/Credentials ---
$tenantId      = Get-AutomationVariable -Name "TenantID"
$credentials   = Get-AutomationPSCredential -Name 'D365Offboarding'
$clientId      = $credentials.GetNetworkCredential().Username
$clientSecret  = $credentials.GetNetworkCredential().Password
$foBaseUrl     = "" # D365 Environment url
$odataEntity   = "SystemUsers"

$workspaceId             = ""
$WorkspaceSubscriptionId = ""

$d365LicenseGroupIds = @(
    "", # D365-License-Supplychain management
    "",  # D365-License-FinanceAttach
    "",   # D365-License-Operations Activity
    "" # D365-License-Restricted Users
)

$d365LicenseGroupNames = @{
    "" = "D365-License-Supplychain management"
    "" = "D365-License-FinanceAttach"
    "" = "D365-License-Operations Activity"
    "" = "D365-License-Restricted Users"
}


function DisableUsers {
    param(
        [Parameter(Mandatory)]
        [array]$UsersToDisable,

        [Parameter(Mandatory)]
        [string]$FoBaseUrl,

        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter(Mandatory)]
        [string[]]$LicenseGroupIds,

        [Parameter(Mandatory)]
        [hashtable]$LicenseGroupNames
    )

    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
        "Accept"        = "application/json"
    }

    foreach ($user in $UsersToDisable) {

        $removedGroups = @()
  
        Write-Output "Processing: $($user.UserId) ($($user.Email))"

        # --- Step 1: Disable in D365 ---
        $foUrl = "$FoBaseUrl/data/SystemUsers(UserID='$($user.UserId)')"
        $body  = @{ Enabled = $false } | ConvertTo-Json

        try {
            Invoke-RestMethod -Method Patch -Uri $foUrl -Headers $headers -Body $body -ErrorAction Stop
            Write-Output "[OK] Disabled in D365: $($user.UserId)"
        } catch {
            Write-Output "[FAILED] Could not disable in D365: $($user.UserId) — $_"
        }

        # --- Step 2: Remove from D365 license groups ---
        $entraObjectId = $null
        try {
            $entraObjectId = (Get-MgUser -UserId $user.Email -Property "Id" -ErrorAction Stop).Id
        } catch {
            Write-Output "[WARN] Could not resolve Entra Object ID for $($user.Email) — skipping group removal"
        }

        if ($entraObjectId) {
            foreach ($groupId in $LicenseGroupIds) {

                # Check membership before attempting removal to avoid noisy errors
                $isMember = $false
                try {
                    $memberCheck = Get-MgGroupMember -GroupId $groupId -All -ErrorAction Stop |
                                   Where-Object { $_.Id -eq $entraObjectId }
                    $isMember = $null -ne $memberCheck
                } catch {
                    Write-Output "[WARN] Could not check membership for group $groupId — $_"
                }

                if ($isMember) {
                    try {
                        Remove-MgGroupMemberByRef -GroupId $groupId -DirectoryObjectId $entraObjectId -ErrorAction Stop
                        $removedGroups += $LicenseGroupIds[$groupId]
                        Write-Output "[OK] Removed from group $groupId"
                    } catch {
                        Write-Output "[FAILED] Could not remove from group $groupId — $_"
                    }
                } else {
                    Write-Output "[SKIP] Not a member of group $groupId"
                }
            }
        }

        $user.LicensesRemoved = ($removedGroups -join "; ")

        # --- Step 3: Resolve manager and send notification email ---
        $manager = $null
        try {
            $managerId = Get-MgUserManager -UserId $user.Email -ErrorAction SilentlyContinue
            if ($managerId.Id) {
                $manager = Get-MgUser -UserId $managerId.Id -Property "Mail,DisplayName" -ErrorAction SilentlyContinue
            }
        } catch {}

        $ccRecipients = @()
        if ($manager.Mail) {
            $ccRecipients += @{ emailAddress = @{ address = $manager.Mail } }
            Write-Output "[INFO] Manager found: $($manager.Mail) — will be CC'd"
        } else {
            Write-Output "[INFO] No manager found — sending without CC"
        }

        $firstName = ($user.UserId.Split(".")[0])
        if ($firstName.Length -gt 0) {
            $firstName = $firstName.Substring(0,1).ToUpper() + $firstName.Substring(1)
        }

        $mailBody = "
Hi $firstName,

Your Dynamics 365 license has been removed because it hasn't been used for 60 days. Disabling unused licenses helps us reduce costs and minimize security risks.

A branch user license costs 307-428 SEK/month
A back-office user license costs 1290-1480 SEK/month
These license costs are charged directly to your department once per quarter through the IT Fee.

If you need access again, please contact the Service Desk. Please note that the license will be removed again if unused for another 60 days.

Only request a license if you use Dynamics regularly - for occasional needs, it's more efficient to ask a colleague for reports or information.

Best regards,
Group IT
"

        $emailPayload = @{
            message = @{
                subject      = "Your D365 Account Has Been Disabled"
                body         = @{
                    contentType = "Text"
                    content     = $mailBody
                }
                toRecipients = @(
                    @{ emailAddress = @{ address = $user.Email } }

                )
                ccRecipients = $ccRecipients
            }
            saveToSentItems = "true"
        }

        try {
            Send-MgUserMail -UserId "email@email.com" -BodyParameter $emailPayload
            Write-Output "[OK] Notification sent to $($user.Email)"
        } catch {
            Write-Output "[FAILED] Could not send email to $($user.Email) — $_"
        }
    }
}


Connect-MgGraph -Identity
Connect-AzAccount -Identity
Set-AzContext -SubscriptionId $WorkspaceSubscriptionId

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
$odataUrl = "$foBaseUrl/data/$odataEntity/?`$filter=Enabled eq true"
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

# --- Query Sentinel ---
Write-Output "Querying Sentinel for sign-in history..."

$query = @"
SigninLogs
| where AppDisplayName == 'Microsoft Dynamics ERP' or AppDisplayName == 'D365 PROD Tasklet Integration'
| summarize LastSignIn = max(TimeGenerated), AppLastUsed = arg_max(TimeGenerated, AppDisplayName) by UserPrincipalName
"@

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

# --- Fetch D365 Role Assignments ---
$roleUrl = "$foBaseUrl/data/SecurityUserRoles"
try {
    $roleResponse = Invoke-RestMethod -Method Get -Uri $roleUrl -Headers $headers
    $allRoles     = $roleResponse.value
    Write-Output "Retrieved $($allRoles.Count) role assignments"
} catch {
    Write-Output "Failed to get role assignments: $_"
}

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

# --- Parse Sentinel Response ---
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
$cutOff            = (Get-Date).AddDays(-60)
$InactiveUsers     = @()
$ManualReviewUsers = @()

foreach ($user in $filteredUsers) {

    if ([string]::IsNullOrWhiteSpace($user.Email)) {
        Write-Output "Skipping $($user.UserID) - no email"
        continue
    }

    if ($usersToSkip -contains $user.UserID.ToLower()) {
        Write-Output "Skipping $($user.UserID) - WMS role"
        continue
    }

    $email       = $user.Email.ToLower()
    $entraUser   = Get-MgUser -UserId $email -Property "createdDateTime,userPrincipalName" -ErrorAction SilentlyContinue
    $swedishDate = "N/A"
    $parsedDate  = $null

    if ($entraUser.CreatedDateTime) {
        $parsedDate  = [datetime]$entraUser.CreatedDateTime
        $swedishDate = $parsedDate.ToString("yyyy-MM-dd")
    }

    if (($email -notmatch "hydroscand") -or ($user.UserID -match "svc") -or ($parsedDate -ne $null -and $parsedDate -gt $cutOff)) {
        Write-Output "Skipping $($user.UserID)"
        continue
    }

    if (-not $entraUser) {
        Write-Output "Trying displayName match for $($user.UserName)"
        $entraUser = Get-MgUser -Filter "displayName eq '$($user.UserName)'" -All -ErrorAction SilentlyContinue | Select-Object -First 1
    }

    $lastSignInDynamics = $dynamicsLookup[$entraUser.userPrincipalName]
    $lastSignInTasklet  = $taskletLookup[$entraUser.userPrincipalName]

    if ($lastSignInTasklet -ne $null -and $lastSignInTasklet -ge $cutOff) {
        Write-Output "Skipping $($user.UserID) - active in Tasklet"
        continue
    }

    if (-not $entraUser) {
        $ManualReviewUsers += [pscustomobject]@{
            UserId             = $user.UserID
            Email              = $user.Email
            LastSignIn         = "N/A"
            AppLastUsed        = "N/A"
            AccountCreatedDate = "N/A"
            "Enabled in D365"  = $user.Enabled
            Reason             = "User Not Found in Entra - Manual Review Required"
        }
    } elseif ($null -eq $lastSignInDynamics -or $lastSignInDynamics -lt $cutOff) {
        $InactiveUsers += [pscustomobject]@{
            UserId             = $user.UserID
            Email              = $entraUser.userPrincipalName
            LastSignIn         = if ($lastSignInDynamics) { $lastSignInDynamics.ToString("yyyy-MM-dd") } else { "No record in Sentinel" }
            AppLastUsed        = "Microsoft Dynamics ERP"
            AccountCreatedDate = $swedishDate
            "Enabled in D365"  = $user.Enabled
            Reason             = "Inactive > 60 Days in Dynamics"
            LicensesRemoved    = ""
        }
    }
}

Write-Output "Inactive users (will be auto-disabled): $($InactiveUsers.Count)"
Write-Output "Manual review users (not in Entra):     $($ManualReviewUsers.Count)"

# --- Disable Inactive Users ---
if ($InactiveUsers.Count -gt 0) {
    DisableUsers `
        -UsersToDisable  $InactiveUsers `
        -FoBaseUrl       $foBaseUrl `
        -AccessToken     $accessToken `
        -LicenseGroupIds $d365LicenseGroupIds `
        -LicenseGroupNames $d365LicenseGroupNames
} else {
    Write-Output "No inactive users to disable."
}

# --- Build Reporting CSV (all flagged users) ---
$AllFlaggedUsers = $InactiveUsers + $ManualReviewUsers

try {
    $CsvContent = $AllFlaggedUsers | ConvertTo-Csv -NoTypeInformation
    $CsvBytes   = [System.Text.Encoding]::UTF8.GetBytes($CsvContent -join "`n")
    $CsvBase64  = [System.Convert]::ToBase64String($CsvBytes)
} catch {
    Write-Error "Failed to convert data to CSV: $_"
    exit 1
}

# --- Email 1: Reporting email with full CSV ---
$reportingPayload = @{
    message = @{
        subject = "Inactive Users found in PROD"
        body    = @{
            contentType = "Text"
            content     = "Please find attached the inactive D365 users report for PROD.`n`nAuto-disabled: $($InactiveUsers.Count) users`nRequire manual review (not in Entra): $($ManualReviewUsers.Count) users"
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

try {
    Send-MgUserMail -UserId "email@domain.com" -BodyParameter $reportingPayload
    Write-Output "Reporting email sent."
} catch {
    Write-Output "Failed to send reporting email: $_"
}

# --- Email 2: Service desk email for manual review users ---
if ($ManualReviewUsers.Count -gt 0) {

    $tableRows = ($ManualReviewUsers | ForEach-Object {
        "<tr>
            <td>$($_.UserId)</td>
            <td>$($_.Email)</td>
        </tr>"
    }) -join "`n"

    $htmlTable = @"
<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>
    <thead>
        <tr>
            <th>UserID</th>
            <th>Email on record</th>
        </tr>
    </thead>
    <tbody>
        $tableRows
    </tbody>
</table>
"@

    $serviceDeskBody = @"

The automated D365 inactive user runbook found <b>$($ManualReviewUsers.Count)</b> user(s) that are enabled in D365 but could not be matched to an account in Entra ID.<br><br>

These users have <b>NOT</b> been automatically disabled and require manual investigation:<br><br>

$htmlTable

<br><br>
Please check each user and take appropriate action (disable in D365, remove from groups, or confirm if they should remain active).<br><br>

"@

}

$serviceDeskPayload = @{
    message = @{
        subject = "D365 Inactive Users - Manual Review ($($ManualReviewUsers.Count) users)"
        body    = @{
            contentType = "html"
            content     = $serviceDeskBody
        }
        toRecipients = @(
            @{ emailAddress = @{ address = "email@domain.com" } }
        )
    }
    saveToSentItems = "true"
}

try {
    Send-MgUserMail -UserId "email@domain.com" -BodyParameter $serviceDeskPayload
    Write-Output "Service desk email sent for $($ManualReviewUsers.Count) manual review users."
} catch {
    Write-Output "Failed to send service desk email: $_"
}

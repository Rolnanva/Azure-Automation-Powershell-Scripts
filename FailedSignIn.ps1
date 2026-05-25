
# Tenant & App Information
$tenantId = ""
$clientId     = Get-AutomationVariable -Name "pbiGraphId"
$clientSecret = Get-AutomationVariable -Name "pbiGraphSecret"

# OAuth Token URL
$tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Token request body
$body = @{
    client_id     = $clientId
    client_secret = $clientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
}

# Get token
$tokenResponse = Invoke-RestMethod -Method POST -Uri $tokenUrl -Body $body
$accessToken   = $tokenResponse.access_token

# Headers
$headers = @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type"  = "application/json"
}

# Last 24 hours
$startTime = (Get-Date).AddDays(-1).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$endTime   = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

$logUrl = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=createdDateTime ge $startTime and createdDateTime le $endTime&`$top=999"

$logsResponse = Invoke-RestMethod -Uri $logUrl -Headers $headers -Method GET
$logs = $logsResponse.value

# Group ALL sign-ins (total attempts) by AppDisplayName + AppId
$totalUserApp = $logs |
    Group-Object -Property AppDisplayName, AppId |
    Select-Object @{
        Name = 'AppDisplayName'
        Expression = { $_.Group[0].AppDisplayName }
    }, @{
        Name = 'AppId'
        Expression = { $_.Group[0].AppId }
    }, @{
        Name = 'TotalCount'
        Expression = { $_.Count }
    }

Write-Output "TotalLogs"
Write-Output $totalUserApp

$failedLogs = $logs | Where-Object { $_.Status.ErrorCode -ne 0 }

# Group FAILED sign-ins
$failedUserApp = $failedLogs |
    Group-Object -Property AppDisplayName, AppId |
    Select-Object @{
        Name = 'AppDisplayName'
        Expression = { $_.Group[0].AppDisplayName }
    }, @{
        Name = 'AppId'
        Expression = { $_.Group[0].AppId }
    }, @{
        Name = 'FailCount'
        Expression = { $_.Count }
    }
Write-Output "Failed Logs:"
Write-Output $failedUserApp 
# Join failed attempts with total attempts (matching AppId + AppName)
$mergedData = foreach ($fail in $failedUserApp) {
    $total = $totalUserApp | Where-Object {
        $_.AppId -eq $fail.AppId -and $_.AppDisplayName -eq $fail.AppDisplayName
    }

    if ($total) {
        [PSCustomObject]@{
            AppDisplayName = $fail.AppDisplayName
            AppId          = $fail.AppId
            FailCount      = $fail.FailCount
            TotalCount     = $total.TotalCount
        }
    }
}

# Only alert if failCount >= 3
$alertApps = $mergedData | Where-Object { $_.FailCount -ge 3 }
if ($alertApps.Count -gt 0) {

    # Build HTML Table
    $htmlTable = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;'>
<tr style='background-color:#f2f2f2;'>
<th>Application</th>
<th>App ID</th>
<th>Failure Count</th>
</tr>"

foreach ($entry in $alertApps) {
    $htmlTable += "<tr>
<td>$($entry.AppDisplayName)</td>
<td>$($entry.AppId)</td>
<td style='text-align:center;'>$($entry.FailCount) / $($entry.TotalCount)</td>
</tr>"
}

    $htmlTable += "</table>"

    $htmlBody = "<p>The following apps had 3 or more failed sign-ins in the last 24 hours:</p>$htmlTable"

    # Graph sendMail endpointarning
    $sendMailUrl = "https://graph.microsoft.com/v1.0/users/it@hydroscandgroup.com/sendMail"

    $emailPayload = @{
        message = @{
            subject = "Alert: Multiple Failed Sign-ins Detected"
            body = @{
                contentType = "HTML"
                content = $htmlBody
            }
            toRecipients = @(
                @{ emailAddress = @{ address = "email@domain.com" } }

            )
        }
        saveToSentItems = $true
    } | ConvertTo-Json -Depth 10

    Invoke-RestMethod -Uri $sendMailUrl -Headers $headers -Method POST -Body $emailPayload -ContentType "application/json"

    Write-Output "HTML alert email sent."
}
else {
    Write-Output "No applications met the failed sign-in threshold."
}

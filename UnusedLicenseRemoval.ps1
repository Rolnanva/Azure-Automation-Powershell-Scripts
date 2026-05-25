
$SenderEmail  = "email@domain.com"
$ReportPeriod = "D90"
$ReportDate   = (Get-Date).ToString("yyyy-MM-dd")

$CopilotSkuId = ""

# Thresholds
$ReminderDays = 35
$RemovalDays  = 65

$ActionLog = @()

Write-Output "Connecting to Microsoft Graph..."
Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop



Write-Output "Retrieving users..."

$AllUsers = Get-MgUser -All -Property "DisplayName,UserPrincipalName,AssignedLicenses"

$CopilotUsers = $AllUsers | Where-Object {
    $_.AssignedLicenses.SkuId -contains $CopilotSkuId
}


Write-Output "Total users: $($AllUsers.Count)"
Write-Output "Copilot users: $($CopilotUsers.Count)"



Write-Output "Fetching Copilot usage report..."

$ReportUri = "https://graph.microsoft.com/v1.0/copilot/reports/" +
             "microsoft.graph.getMicrosoft365CopilotUsageUserDetail(period='$ReportPeriod')"

$CsvString = (Invoke-MgGraphRequest -Uri $ReportUri -Method GET -OutputType HttpResponseMessage).Content.ReadAsStringAsync().Result
$CsvString = $CsvString -replace "`r`n", "`n" -replace "`r", "`n"

$ActivityRows = $CsvString | ConvertFrom-Csv

$Lookup = @{}
foreach ($r in $ActivityRows) {
    if ($r.'User Principal Name') {
        $Lookup[$r.'User Principal Name'.ToLower()] = $r
    }
}



Write-Output "Processing user activity..."

$Today = Get-Date

$ProcessedUsers = foreach ($User in $CopilotUsers) {

    $Activity = $Lookup[$User.UserPrincipalName.ToLower()]

    if ($Activity -and $Activity.'Last Activity Date') {
        try {
            $LastDate = [datetime]$Activity.'Last Activity Date'
            $DaysInactive = ($Today - $LastDate).Days
        }
        catch {
            $LastDate = $null
            $DaysInactive = 999
        }
    }
    else {
        $LastDate = $null
        $DaysInactive = 999
    }

    [PSCustomObject]@{
        DisplayName       = $User.DisplayName
        UserPrincipalName = $User.UserPrincipalName
        LastActivityDate  = $LastDate
        DaysInactive      = $DaysInactive
    }
}

Write-Output "Users processed: $($ProcessedUsers.Count)"



$ReminderUsers = $ProcessedUsers | Where-Object {
    $_.DaysInactive -ge $ReminderDays -and $_.DaysInactive -lt $RemovalDays
}

$RemoveUsers = $ProcessedUsers | Where-Object {
    $_.DaysInactive -ge $RemovalDays
}

$NoActionUsers = $ProcessedUsers | Where-Object {
    $_.DaysInactive -lt $ReminderDays
}

Write-Output "No Action users: $($NoActionUsers.Count)"
Write-Output "Reminder users: $($ReminderUsers.Count)"
Write-Output "Removal users: $($RemoveUsers.Count)"


foreach ($User in $NoActionUsers) {

    Write-Output "[NO ACTION] $($User.UserPrincipalName) - $($User.DaysInactive) days"

    $ActionLog += [PSCustomObject]@{
        UserPrincipalName = $User.UserPrincipalName
        DisplayName       = $User.DisplayName
        Action            = "No Action"
        DaysInactive      = $User.DaysInactive
        LastActivityDate  = $User.LastActivityDate
        License           = "Copilot"
        ReportDate        = $ReportDate
    }
}

# REMINDERS
foreach ($User in $ReminderUsers) {

    Write-Output "[REMIND] $($User.UserPrincipalName) - $($User.DaysInactive) days"

    $ActionLog += [PSCustomObject]@{
        UserPrincipalName = $User.UserPrincipalName
        DisplayName       = $User.DisplayName
        Action            = "Reminded"
        DaysInactive      = $User.DaysInactive
        LastActivityDate  = $User.LastActivityDate
        License           = "Copilot"
        ReportDate        = $ReportDate
    }

    $Body = @"
Hello $($User.DisplayName),

Our records show that you have one or more licenses assigned that has not been used for the past 30 days.

The following license(s) are flagged as inactive:
- Copilot

If the license(s) remains unused for the next 30 days, it will be automatically removed.

If you wish to keep the license, no contact with IT is required — simply sign in and use the service within the next 30 days.

Kind regards,  
IT Department
"@

    Send-MgUserMail -UserId $SenderEmail -Message @{
        Subject = "Inactive License – Action Required Within 30 Days"
        Body = @{
            ContentType = "Text"
            Content     = $Body
        }
        ToRecipients = @(
            @{ EmailAddress = @{ Address = $User.UserPrincipalName } }

        )
    } -SaveToSentItems:$false
}

# REMOVALS
foreach ($User in $RemoveUsers) {

    Write-Output "[REMOVE] $($User.UserPrincipalName) - $($User.DaysInactive) days"

    $Licenses = (Get-MgUser -UserId $User.UserPrincipalName -Property AssignedLicenses).AssignedLicenses

    if ($Licenses.SkuId -notcontains $CopilotSkuId) {

        Write-Output "[SKIPPED] $($User.UserPrincipalName) - license missing"

        $ActionLog += [PSCustomObject]@{
            UserPrincipalName = $User.UserPrincipalName
            DisplayName       = $User.DisplayName
            Action            = "Skipped"
            DaysInactive      = $User.DaysInactive
            LastActivityDate  = $User.LastActivityDate
            License           = "Copilot"
            ReportDate        = $ReportDate
        }

        continue
    }

    try {
    Write-Output "[REMOVE] Attempting license removal for $($User.UserPrincipalName)"

    Set-MgUserLicense -UserId $User.UserPrincipalName `
        -RemoveLicenses @($CopilotSkuId) `
        -AddLicenses @() `
        -ErrorAction Stop `

    Write-Output "[SUCCESS] License removal succeeded for $($User.UserPrincipalName)"

    $ActionLog += [PSCustomObject]@{
        UserPrincipalName = $User.UserPrincipalName
        DisplayName       = $User.DisplayName
        Action            = "Removed (WhatIf)"
        DaysInactive      = $User.DaysInactive
        LastActivityDate  = $User.LastActivityDate
        License           = "Copilot"
        ReportDate        = $ReportDate
        Status            = "Success"
        ErrorMessage      = ""
    }
}
catch {
    $ErrorMessage = $_.Exception.Message

    Write-Output "[ERROR] Failed to remove license for $($User.UserPrincipalName)"
    Write-Output "[ERROR DETAILS] $ErrorMessage"

    $ActionLog += [PSCustomObject]@{
        UserPrincipalName = $User.UserPrincipalName
        DisplayName       = $User.DisplayName
        Action            = "Removal Failed"
        DaysInactive      = $User.DaysInactive
        LastActivityDate  = $User.LastActivityDate
        License           = "Copilot"
        ReportDate        = $ReportDate
        Status            = "Failed"
        ErrorMessage      = $ErrorMessage
    }

    continue
}
    $Body = @"
Hi $($User.DisplayName),

Your Copilot license is scheduled for removal due to inactivity ($($User.DaysInactive) days).

If needed, contact IT.

Thanks,  
IT
"@

    Send-MgUserMail -UserId $SenderEmail -Message @{
        Subject = "Copilot License Removal Notice"
        Body = @{
            ContentType = "Text"
            Content     = $Body
        }
        ToRecipients = @(
            @{ EmailAddress = @{ Address = $User.UserPrincipalName } }

        )
    } -SaveToSentItems:$false
}



Write-Output "Generating CSV report..."

$CsvPath = "Copilot_ActionLog_$ReportDate.csv"

$ActionLog | Sort-Object Action, DaysInactive |
    Export-Csv $CsvPath -NoTypeInformation -Encoding UTF8

Write-Output "CSV saved: $CsvPath"

$FileBytes = [System.IO.File]::ReadAllBytes($CsvPath)
$FileBase64 = [System.Convert]::ToBase64String($FileBytes)

$AdminBody = @"
Copilot License Activity Report - $ReportDate

Total users processed: $($ActionLog.Count)

See attached CSV for full details.
"@

Write-Output "Sending admin report..."

Send-MgUserMail -UserId $SenderEmail -Message @{
    Subject = "Copilot License Automation Report - $ReportDate"
    Body = @{
        ContentType = "Text"
        Content     = $AdminBody
    }
    ToRecipients = @(
        @{ EmailAddress = @{ Address = "email@domain.com" } }
    )
    Attachments = @(
        @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            Name          = "Copilot_ActionLog_$ReportDate.csv"
            ContentType   = "text/csv"
            ContentBytes  = $FileBase64
        }
    )
} -SaveToSentItems:$false

Write-Output "Script completed successfully."
Write-Output "Total logged actions: $($ActionLog.Count)"

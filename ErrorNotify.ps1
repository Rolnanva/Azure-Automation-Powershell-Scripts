
$AutomationAccount = ""
$ResourceGroup = ""
$SenderUPN = ""
$To = ""
$Subject = "Automation Runbook Errors/Warnings (Last 48h)"

# Authenticate using Azure Automation Managed Identity
Connect-AzAccount -Identity

# Connect to Microsoft Graph using Managed Identity
Connect-MgGraph -Identity

# Get jobs from last 48 hours
$cutoff = (Get-Date).AddHours(-48)

$jobs = Get-AzAutomationJob -ResourceGroupName $ResourceGroup -AutomationAccountName $AutomationAccount |
    Where-Object { $_.StartTime -ge $cutoff }

$alerts = @()

foreach ($job in $jobs) {

    $errors = Get-AzAutomationJobOutput `
        -ResourceGroupName $ResourceGroup `
        -AutomationAccountName $AutomationAccount `
        -Id $job.JobId `
        -Stream Error

    $warnings = Get-AzAutomationJobOutput `
        -ResourceGroupName "$ResourceGroup" `
        -AutomationAccountName $AutomationAccount `
        -Id $job.JobId `
        -Stream Warning

    if ($errors.Count -gt 0 -or $warnings.Count -gt 0) {
        $alerts += [PSCustomObject]@{
            Runbook  = $job.RunbookName
            Status   = $job.Status
            Errors   = $errors.Count
            Warnings = $warnings.Count
            Started  = $job.StartTime
        }
    }
}

# If nothing has issues → stop here (no email)
if ($alerts.Count -eq 0) {
    Write-Output "No errors or warnings found. No email sent."
    return
}


$grouped = $alerts | Group-Object Runbook


# Build the plain text email body
$Body = "Jobs with errors/warnings in last 48 hours:`n`n"

foreach ($group in $grouped) {
    $Body += ($group.Group | Format-Table Runbook, Status, Errors, Warnings, Started -AutoSize | Out-String)
    $Body += "`n"
}


# Build the Graph message object
$message = @{
    Subject = $Subject
    Body = @{
        ContentType = "Text"
        Content     = $Body
    }
    ToRecipients = @(
        @{ EmailAddress = @{ Address = $To } }
    )
}

# Send the email using Microsoft Graph
Send-MgUserMail -UserId $SenderUPN -Message $message -SaveToSentItems


foreach ($group in $grouped) {
    $group.Group | Format-Table Runbook, Status, Errors, Warnings, Started -AutoSize
}


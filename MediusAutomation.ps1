
Connect-MgGraph -Identity

#App ID for medius enterprise app
$appId = ""

$cutOff = (Get-Date).AddDays(-1).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

Write-Output "Fetching groups assigned to Enterprise Application"

# Retrieve the Service Principal (Enterprise App) by App ID
$servicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$appId'"

if (-not $servicePrincipal) {
    Write-Output " Enterprise Application with App ID '$appId' not found!"
    exit
}

# Get all groups assigned to this Enterprise Application
$groups = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $servicePrincipal.Id | 
          Where-Object { $_.PrincipalType -eq "Group" } |
          Select-Object @{Name="Name"; Expression={ $_.PrincipalDisplayName }},
                        @{Name="Id"; Expression={ $_.PrincipalId }}

if (-not $groups) {
    Write-Output " No groups found assigned to the application!"
    exit
}

Write-Output "Groups in enterprise app"
Write-Output $groups.Name

foreach ($group in $groups){

    $auditLogs = Get-MgAuditLogDirectoryAudit -Filter "activityDisplayName eq 'Add member to group' and targetResources/any(tr: tr/id eq '$($group.Id)') and activityDateTime ge $cutOff" -All

    if (-not $auditLogs) {
        Write-Output "No new accounts added to $($group.Name) in the last 24 hours"
    }

        Write-Output "`nNew account(s) added to $($group.Name) in the last 24 hours:"
        
        $auditLogs | ForEach-Object {
            $addedUserUpn = ($_.targetResources | Where-Object { $_.type -eq "User" }).userPrincipalName
            $groupName = $group.Name
            $dateAdded = $_.activityDateTime
            [PSCustomObject]@{
                DateAdded    = $dateAdded
                AddedUserUPN = $addedUserUpn
                GroupName    = $groupName
            }| Format-Table -AutoSize

        $body = "
Hello,

A new member has been added to the group $groupName.

Details:
- User: $addedUserUpn
- Group: $groupName
- Date Added: $dateAdded

Please review this change.


"

        $emailPayload = @{
            message = @{
                subject = "New user added to $($group.Name)"
                body = @{
                    contentType = "text"
                    content = $body
                }
                toRecipients = @(
                    @{ emailAddress = @{ address = "email@domain.com" } }


                )
            }
            saveToSentItems = $true
        }

            Send-MgUserMail -UserId "email@domain.com" -BodyParameter $emailPayload
            Write-Output "Email sent"

        if ($group.Name -eq "Sweden - Medius Users" -or $group.Name -eq "Denmark - Medius Users"){
            $body = "
Hello,

A new member has been added to the group $groupName.

Details:
- User: $addedUserUpn
- Group: $groupName
- Date Added: $dateAdded

Please review this change.


"

        $emailPayload = @{
            message = @{
                subject = "New user added to $($group.Name)"
                body = @{
                    contentType = "text"
                    content = $body
                }
                toRecipients = @(
                    @{ emailAddress = @{ address = "email@domain.se" } }


                )
            }
            saveToSentItems = $true
        }

            Send-MgUserMail -UserId "email@domain.com" -BodyParameter $emailPayload
            Write-Output "Email sent"              
        }

        }    
    }
    


foreach ($group in $groups){

    Write-Output "`nChecking offboarding for group: $($group.Name)"

    # Get users removed from group
    $auditLogsRemoved = Get-MgAuditLogDirectoryAudit -Filter "activityDisplayName eq 'Remove member from group' and targetResources/any(tr: tr/id eq '$($group.Id)') and activityDateTime ge $cutOff" -All

    if (-not $auditLogsRemoved) {
        Write-Output "No users removed from $($group.Name) in the last 7 days"
        continue
    }

    Write-Output "`nUser(s) removed from $($group.Name) in the last 7 days:"

    foreach ($log in $auditLogsRemoved) {

        $removedUserUpn = ($log.targetResources | Where-Object { $_.type -eq "User" }).userPrincipalName
        $groupName = $group.Name
        $dateRemoved = $log.activityDateTime

        [PSCustomObject]@{
            DateRemoved   = $dateRemoved
            RemovedUserUPN = $removedUserUpn
            GroupName     = $groupName
        } | Format-Table -AutoSize

        # Email body
        $body = "
Hello,

A user has been removed from the group $groupName.

Details:
- User: $removedUserUpn
- Group: $groupName
- Date Removed: $dateRemoved

Please review this change.
"

        $emailPayload = @{
            message = @{
                subject = "User removed from $($group.Name)"
                body = @{
                    contentType = "text"
                    content = $body
                }
                toRecipients = @(
                    @{ emailAddress = @{ address = "email@domain.com" } }

                )
            }
            saveToSentItems = $true
        }

        Send-MgUserMail -UserId "email@domain.com" -BodyParameter $emailPayload
        Write-Output "Offboarding email sent"


        # Special handling for SE/DK Medius groups
        if ($group.Name -eq "Sweden - Medius Users" -or $group.Name -eq "Denmark - Medius Users") {

            $emailPayload = @{
                message = @{
                    subject = "User removed from $($group.Name)"
                    body = @{
                        contentType = "text"
                        content = $body
                    }
                    toRecipients = @(
                        @{ emailAddress = @{ address = "email@domain.se" } }

                    )
                }
                saveToSentItems = $true
            }

            Send-MgUserMail -UserId "email@domain.com" -BodyParameter $emailPayload
            Write-Output "Offboarding email sent (invoice)"
        }
    }
}
       


       


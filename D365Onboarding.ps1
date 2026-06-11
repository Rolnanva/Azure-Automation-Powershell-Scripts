
#Tenant Id
$tenantId       = get-automationvariable -name "TenantID"

#Fin&Ops Credentials
$credentials= Get-AutomationPSCredential -Name 'D365 UAT Integration Cred'
$clientId = $credentials.GetNetworkCredential().Username
$clientSecret = $credentials.GetNetworkCredential().Password
$foBaseUrl      = ""

#FreshService Credentials
$FreshApiKey =  get-automationvariable -name "FreshApiKey"
$FreshBaseUrl = ""

#Connect to Graph
Connect-MgGraph -Identity

#Get access token for D365
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

$ErrorList = New-Object System.Collections.ArrayList


function CreateD365User{

    param(
        [Parameter(mandatory)]
        [string]$UserEmail,

        [Parameter(Mandatory)]
        [string]$Company,

        [Parameter(mandatory)]
        [Array]$WarehouseIds

    )

    $user = $entrausers | Where-Object {$_.UserPrincipalName -eq "$UserEmail"}

    $UserId = "$($user.GivenName).$($user.Surname)"

    #Create D365 User
    $CreateUser = "$foBaseUrl/data/SystemUsers"
    $headers = @{
        "Authorization" = "Bearer $accessToken"
        "Accept"        = "application/json"
        "Content-Type" = "application/json"
    }
    $CreateUserBody = @{
        "NetworkDomain" = "https://sts.windows.net/"
        "Helplanguage" = "en-us"
        "UserID" = $UserId
        "UserInfo_language" = "en-us"
        "UserName" = $user.displayName
        "Alias" = $userEmail
        "Email" = $userEmail
        "AccountType" = "ClaimsUser"
        "Company" = $Company
        "Enabled" = $true
    } | ConvertTo-Json -Depth 3

    try {
        Write-Output "Trying to Create user $UserId"
        Invoke-RestMethod -Method Post -Uri $CreateUser -Headers $headers -Body $CreateUserBody
        Write-Output "User $UserId Created"
    } catch {
        Write-Output "Failed to create user in D365: $_"
        try {
            $parsed = $_.ErrorDetails.Message | ConvertFrom-Json
            $cleanMessage = $parsed.error.innererror.message
        }catch {
            $cleanMessage = $_.Exception.Message
        }

        $ErrorList.Add([PSCustomObject]@{
            Step = "Create User"
            Message = $cleanMessage
        })

    }

    #Create D365 Worker
    $StartDate = (Get-Date).AddDays(-1).ToUniversalTime().ToString("yyyy-MM-ddT00:00:00Z")
    $CreateWorker = "$foBaseUrl/data/Workers"
    $CreateWorkerBody = @{
        "PersonnelNumber" = $user.onPremisesSamAccountName
        "FirstName" = $user.GivenName
        "LastName" = $user.Surname
        "OriginalHireDateTime" = $StartDate
        "PrimaryContactEmail" = $userEmail
        "PrimaryContactEmailDescription" = "Email"
    } | ConvertTo-Json -Depth 3

    
    try{
        Write-Output "Trying to Create Worker for $UserId"
        Invoke-RestMethod -Method Post -Uri $CreateWorker -Headers $headers -Body $CreateWorkerBody
        Write-Output "Worker $UserId Created"


    }catch{
        Write-Output "OData request failed Worker Creation: $_"
        try {
            $parsed = $_.ErrorDetails.Message | ConvertFrom-Json
            $cleanMessage = $parsed.error.innererror.message
        }catch {
            $cleanMessage = $_.Exception.Message
        }

        $ErrorList.Add([PSCustomObject]@{
            Step = "Create Worker"
            Message = $cleanMessage
        })

    }

    #Create Employment for User
    $CreateEmployement = "$foBaseUrl/data/Employments"
    $CreateEmployementBody = @{
        "EmploymentStartDate" = $StartDate
        "LegalEntityId" = $Company
        "PersonnelNumber" = $user.onPremisesSamAccountName
        "WorkerType" = "Employee"
    } | ConvertTo-Json -Depth 3

    try{
        Write-Output "Trying to Create Employment for $UserId"
        Invoke-RestMethod -Method Post -Uri $CreateEmployement -Headers $headers -Body $CreateEmployementBody
        Write-Output "Employment for $UserId Created"

    }catch{
        Write-Output "OData request failed Employment Creation: $_"
        try {
            $parsed = $_.ErrorDetails.Message | ConvertFrom-Json
            $cleanMessage = $parsed.error.innererror.message
        }catch {
            $cleanMessage = $_.Exception.Message
        }

        $ErrorList.Add([PSCustomObject]@{
            Step = "Create Employee"
            Message = $cleanMessage
        })

    }

    #Get Party Id for User
    $GetWorker = "$foBaseUrl/data/Workers/?`$filter=PersonnelNumber eq '$($user.onPremisesSamAccountName)'"

    try{
        $WorkerRepsonse = Invoke-RestMethod -Method Get -Uri $GetWorker -Headers $headers
        $Worker = $WorkerRepsonse.value
    }catch{
        Write-Output "OData request failed to Get Worker: $_"
        try {
            $parsed = $_.ErrorDetails.Message | ConvertFrom-Json
            $cleanMessage = $parsed.error.innererror.message
        }catch {
            $cleanMessage = $_.Exception.Message
        }

        $ErrorList.Add([PSCustomObject]@{
            Step = "Get Worker PartyId"
            Message = $cleanMessage
        })

    }

    #Connect Worker to User
    $CreatePersonUser = "$foBaseUrl/data/PersonUsers"
    $CreatePersonUserBody = @{
        "UserId" = $UserId
        "PartyNumber" = $Worker.PartyNumber
        "ValidFrom" = $Worker.OriginalHireDateTime
        "UserEmail" = $userEmail
        "PersonNameAlias" = $user.GivenName
    } | ConvertTo-Json -Depth 3

    try{
        Write-Output "Trying to add $UserId to person"
        Invoke-RestMethod -Method Post -Uri $CreatePersonUser -Headers $headers -Body $CreatePersonUserBody
        Write-Output "Person added for $UserId"
    }catch{
        Write-Output "OData request failed PersonUser Creation: $_"
        try {
            $parsed = $_.ErrorDetails.Message | ConvertFrom-Json
            $cleanMessage = $parsed.error.innererror.message
        }catch {
            $cleanMessage = $_.Exception.Message
        }

        $ErrorList.Add([PSCustomObject]@{
            Step = "Connect Worker to User"
            Message = $cleanMessage
        })

    }


    # If user need access to more than 1 warehouse add them one by one
    foreach($warehouseId in $WarehouseIds){  

        #Get Operationsite from warehouseid
        $GetWarehouse = "$foBaseUrl/data/Warehouses/?`$filter=WarehouseId  eq '$WarehouseId'&cross-company=true"


        try{
            $WarehouseResponse = Invoke-RestMethod -Method Get -Uri $GetWarehouse -Headers $headers
            $Warehouse = $WarehouseResponse.value
        }catch{
            Write-Output "Failed to add user to Warehouse: $_"
        try {
            $parsed = $_.ErrorDetails.Message | ConvertFrom-Json
            $cleanMessage = $parsed.error.innererror.message
        }catch {
            $cleanMessage = $_.Exception.Message
        }

        $ErrorList.Add([PSCustomObject]@{
            Step = "Get Warehouse"
            Message = $cleanMessage
        })

        }

    #Uncomment if usergroups are needed
    <#
        #Get groupid for usergroup
        $GetUsergroup = "$foBaseUrl/data/UserGroups/?`$filter=name eq '$WarehouseName'"

        try{
            $UserGroupResponse = Invoke-RestMethod -Method Get -Uri $GetUsergroup -Headers $headers
            $UserGroupInfo = $UserGroupResponse.value
            Write-Output $UserGroupInfo
        }catch{
            Write-Output "Odata request failed Get usergroup: $_"
        }

        #Add User to UserGroup for warehouse

        $UserGroup = "$foBaseUrl/data/UserGroupUserLists"
        $UserGroupBody = @{
            "groupId" = $UserGroupInfo.GroupId
            "userId" = $UserId
        } | ConvertTo-Json -Depth 3

        try{
            Write-Output "Trying to Add $UserId to UserGroupList"
            Invoke-RestMethod -Method Post -Uri $UserGroup -Headers $headers -Body $UserGroupBody
            Write-Output "$UserId added to UserGroupList"
        }catch{
            Write-Output "OData request failed PersonUser Creation: $_"
        } 
    #>
        #Give User access to Warehouse
        $WarehouseAccess = "$foBaseUrl/data/UserWarehouseAccessLines"

        $WarehouseAccessBody = @{
            "dataAreaId" = $Company
            "UserId" = $UserId
            "InventLocationId" = $WarehouseId
            "InventSiteId" = $Warehouse.OperationalSiteId
        } | ConvertTo-Json -Depth 3

        try{
            Write-Output "Trying to Add User to Warehouse"
            Invoke-RestMethod -Method Post -Uri $WarehouseAccess -Headers $headers -Body $WarehouseAccessBody
            Write-Output "Warehouse added for $UserId"
        }catch{
            Write-Output "OData request failed Adding User to Warehouse: $_"
        try {
            $parsed = $_.ErrorDetails.Message | ConvertFrom-Json
            $cleanMessage = $parsed.error.innererror.message
        }catch {
            $cleanMessage = $_.Exception.Message
        }

        $ErrorList.Add([PSCustomObject]@{
            Step = "Add User to Warehouse"
            Message = $cleanMessage
        })

        }
    }

}

$entrausers = Get-MgUser -All -Select "onPremisesSamAccountName,displayName,UserPrincipalName,GivenName,Surname"


#Get all tickets that contain 'Access for New Employee' & have been approved by manager

$pair = "$($FreshApiKey):X"
$encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))

$headers = @{
    "Authorization" = "Basic $encodedCredentials"
    "Content-Type"  = "application/json"
}

$GetAllTickets = "$FreshBaseUrl/tickets?type=Service Request"



try{
    $Tickets = Invoke-RestMethod -Uri $GetAllTickets -Headers $headers -Method Get
    $AllTickets = $Tickets.tickets
    
}catch{
    Write-Output "Failed to get ticket"
    Write-Output "Error: $($_.Exception.Message)"
}


$FilteredTickets = $AllTickets | Where-Object {
    $_.subject -match "D365 Access for New Employee" -and $_.status -eq 2
}


foreach ($ticket in $FilteredTickets){
    
    $ticketId = $ticket.id

    $GetRequestedItems = "$FreshBaseUrl/tickets/$ticketId/requested_items"


    try{
        $requestedItemsResponse = Invoke-RestMethod -Method Get -Uri $GetRequestedItems -Headers $headers
        $requestedItems = $requestedItemsResponse.requested_items
    }catch {
        Write-Output "Failed to get requested items $_"
    }

    #Get userId(fresh),company and warehouse name from fresh ticket
    $customfields = $requestedItems.custom_fields

    $onBoardUserId = $customfields.choose_a_user

    $userCompany = $customfields.what_legal_entity_are_you_in

    $warehouseIds = $requestedItems.custom_multi_select_dropdowns |
    Where-Object { $_ } |                 # remove nulls
    ForEach-Object { ($_ -split '\s+')[0] }  # take first part (number)


    #Get users mail from their id
    $GetRequester = "$FreshBaseUrl/requesters/$onBoardUserId"


    try{
        $requesterResponse = Invoke-RestMethod -Method Get -Uri $GetRequester -Headers $headers
        $onBoardUser = $requesterResponse.requester
    }catch {
        Write-Output "Failed to get Requester $_"
    }

    $userMail = $onBoardUser.primary_email

    Write-Output $userCompany
    Write-Output $userWarehouse
    Write-Output $userMail
    

    try{
        CreateD365User -UserEmail $userMail -Company $userCompany -WarehouseId $warehouseIds
        Write-Output "User Created in D365"
    }catch{
        Write-Output "Failed to create user in D365 $_"
    }
    
    Write-Output "Error list" $ErrorList

    if ($ErrorList.Count -ge 1){

        $ErrorReply = "$FreshBaseUrl/tickets/$($ticketId)/notes"
        Write-Output $ErrorReply
        $ErrorTable = $ErrorList | ConvertTo-Html -Fragment

        $ErrorReplyBody = @{
            "body" = "Onboarding automation failed. Please complete manually. Errors: $ErrorTable"
        } | ConvertTo-Json -Depth 3
        try{
            Invoke-RestMethod -Method Post -Uri $ErrorReply -Headers $headers -Body $ErrorReplyBody
            Write-Output "Replied to ticket"
        }catch{
            Write-Output "Failed to reply to ticket"
        }


    }else{

        try{
            $closeTicketUri = "$FreshBaseUrl/tickets/$($ticketId)?bypass_mandatory=true"
            Write-Output $closeTicketUri
            $closeTicketBody = @{
                "status" = 5
            }| ConvertTo-Json -Depth 3
            $closeTicket = Invoke-RestMethod -Method Put -Uri $closeTicketUri -Headers $headers -Body $closeTicketBody
        }catch{
            Write-Output "Failed to close ticket $_"
        }
    }
}









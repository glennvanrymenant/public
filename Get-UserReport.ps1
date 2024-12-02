<#
.SYNOPSIS
    This script generates a full report of all users on a given tenant.

.DESCRIPTION

.NOTES
    File Name      : Get-UserReport.ps1
    Author         : Glenn Van Rymenant 
    Prerequisites   : PowerShell 5.1 or later
                      Microsoft Graph PowerShell SDK (https://aka.ms/graphpowershell)

    Version history:

    2.4 rewritten mailbox details per user
    2.3 added OneDrive for Business details per user
    2.2 added all application extension properties available on user objects
    2.1 added optional CSV input with ObjectId/UPN
    2.0 full review and added template
    ...

.EXAMPLE

    ----

    .\Get-UserReport.ps1 -ExportFolderPath "c:\temp\"

    Generates the report and exports it in CSV format in the provided export folder path.

    ----

    .\Get-UserReport.ps1 -ImportCSVFilePath "c:\temp\import.csv" -ExportFolderPath "c:\temp\"

    Generates the report for the users listed in the import CSV file and exports it in CSV format in the provided export folder path.

    ----

#>

### TODO ###
# Add app assignments
# Add recycle bin objects (users/groups)
# Add extension properties (get application objects > extension properties > if enabled for users > get it)
# Add group memberships
# Add emailaddresses from EXO and the count of messages delivered to them
# check if deleted mailboxes exist in EXO

#region parameters

[cmdletbinding()]
param (																		 
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({Test-Path $_ -PathType 'container'})]
    [string] $ExportFolderPath,
    [Parameter(Mandatory = $false)]
    [ValidateScript({Test-Path $_ -PathType 'leaf'})]
    [string] $ImportCSVFilePath,
    [Parameter(Mandatory = $false)]
    [switch] $IncludeGroupMemberships,
    [Parameter(Mandatory = $false)]
    [switch] $IncludeAppAssignments,
    [Parameter(ParameterSetName = "EXO", Mandatory = $false)]
    [switch] $IncludeRecipientType,
    [Parameter(ParameterSetName = "EXO", Mandatory = $false)]
    [switch] $IncludeMailboxDetails,
    [Parameter(Mandatory = $false)]
    [switch] $IncludeOneDriveDetails
)

#endregion parameters

#region parameter validation
#endregion parameter validation

#region variables

# script
$ScriptVersion = "2.5"
Write-Host "[$(Get-Date -Format "HH:mm:ss")][INFO]: Script version: $($ScriptVersion)"
$ScriptStartDateTime = Get-Date

# construct the export path
$ExportFileName = (Get-Date -Format "yyyy-MM-dd_HH-mm_") + "UserReport.csv"
$ExportPath =  Join-Path -Path $ExportFolderPath -ChildPath $ExportFileName

#endregion variables

#region functions

# invoke the Microsoft Graph API
Function Invoke-MgGraphAPI {
    # version 2.8
    # https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands?view=graph-powershell-1.0#using-invoke-mggraphrequest
    param (
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()]
        [String]$Method,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]
        [String]$Endpoint,
        [Parameter(Mandatory = $false)]
        $Body,
        [Parameter(Mandatory = $false)]
        [switch]$Beta
    )

    # If the -Beta switch is used, set the API version to beta
    $APIVersion = if ($PSBoundParameters.ContainsKey('Beta')) { "beta" } else { "v1.0" }

    # Check if the endpoint already contains the API version
    $URI = if ($Endpoint -match "/v1.0/" -or $Endpoint -match "/beta/") {
        "https://graph.microsoft.com$Endpoint"
    } else {
        # Generalize to adapt any endpoint
        if ($Endpoint -match "^v1.0/") {
            "https://graph.microsoft.com/$Endpoint"
        } elseif ($Endpoint -match "^/") {
            "https://graph.microsoft.com/v1.0$Endpoint"
        } else {
            "https://graph.microsoft.com/v1.0/$Endpoint"
        }
    }

    # Create an empty list to store the data
    $Data = [System.Collections.Generic.List[Object]]@()

    # Set the maximum number of attempts
    $MaxAttempts = 5

    # Set the initial number of attempts 
    $Attempts = 1

    # Set the initial value of the success flag to false
    $Success = $false

    # If the method parameter is not specified, set the method to GET
    if (-not $Method) {
        $Method = "GET"
    }

    while ($Attempts -le $MaxAttempts -and -not $Success) {
        try {
            switch ($Method) {
                # If the method is GET, loop until there are no more pages of data
                "GET" {
                    do {
                        $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI -ContentType "application/json"
                        $Success = $true
                        # If the response contains a value property, add the value property to the data list
                        if ($null -ne $Response.Value) {
                            $Data.AddRange($Response.value) 
                            # If the response contains an @odata.nextLink property, set the URI to the value of the @odata.nextLink property (paging)
                            if ($Response.'@odata.nextLink') {
                                # Set the URI to the value of the @odata.nextLink property
                                $URI = $Response.'@odata.nextlink'
                            } 
                            else {
                                # If there is no subsequent @odata.nextLink property, set the URI to null to exit the loop
                                $URI = $null
                            }              
                        }
                        # If the response does not contain a value property, add the response to the data list and set the URI to null to exit the loop
                        else {
                            $Data.Add($Response)
                            $URI = $null
                        }
                    } until ($null -eq $URI)     
                    return $Data
                }
                "POST" {
                    $Body = $Body | ConvertTo-Json
                    $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI -Body $Body -ContentType "application/json"
                    $Success = $true
                    return $Response
                }
                "PUT" {
                    $Body = $Body | ConvertTo-Json
                    $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI -Body $Body -ContentType "application/json"
                    $Success = $true
                    return $Response
                }
                "PATCH" {
                    $Body = $Body | ConvertTo-Json
                    $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI -Body $Body -ContentType "application/json"
                    $Success = $true
                    return $Response
                }
                "DELETE" {
                    $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI
                    $Success = $true
                    return $Response
                }
            }    
        }
        catch {
            # If the call fails due to rate limiting (throttling), retry the call
            if ($_.Exception.Response.StatusCode -eq '429') {
                Write-Host "Encountered throttling. Retrying...`nAttempt: $($Attempts) of $($MaxAttempts)" -ForegroundColor Yellow
                # If the response contains a Retry-After header, wait for the specified number of seconds
                $RetryAfter = $_.Exception.Response.Headers['Retry-After']
                Start-Sleep -Seconds $RetryAfter
                # Increment the number of attempts
                $Attempts++
            }
            else {
                Write-Host -ForegroundColor Red "StatusCode: $($_.Exception.Response.StatusCode)"
                Write-Host -ForegroundColor Red "Exception message: $($_.Exception.Message)"
                # Exit the loop if the exception is not due to throttling
                $Success = $true
            }
        }
    }

    return $Data

}

#endregion functions

#region authentication

$Scopes = [System.Collections.Generic.List[Object]]@("Directory.Read.All","AuditLog.Read.All")

# if parameter 'IncludeOneDriveDetails' is provided 
if ($PSBoundParameters.ContainsKey('IncludeOneDriveDetails')) {
    $Scopes.Add("Reports.Read.All") 
}

# connect to Graph PowerShell
try {
    Write-Host "[$(Get-Date -Format "HH:mm:ss")][INFO]: Connecting to Graph PowerShell..."
    Connect-MgGraph -Scopes $Scopes -ContextScope Process > $null
    Write-Host -ForegroundColor Green "[$(Get-Date -Format "HH:mm:ss")][SUCCESS]: Successfully connected to Graph PowerShell"
}
catch {
    Throw "Failed to connect to Graph PowerShell"
}

# if any parameters from the 'EXO' parameterset were provided
if ($PSCmdlet.ParameterSetName -eq 'EXO') {
    # connect to Exchange Online PowerShell
    try {
        Write-Host "[$(Get-Date -Format "HH:mm:ss")][INFO]: Connecting to Exchange Online PowerShell..."
        Connect-ExchangeOnline -ShowBanner:$false -SkipLoadingCmdletHelp -SkipLoadingFormatData -CommandName ("Get-Recipient", "Get-MailboxStatistics")
        Write-Host -ForegroundColor Green "[$(Get-Date -Format "HH:mm:ss")][SUCCESS]: Successfully connected to Exchange Online PowerShell"
    }
    catch {
        Throw "Failed to connect to Exchange Online PowerShell"
    }
}

#endregion authentication

#region code

# properties not included
# aboutMe, birthDay, hireDate, interests, mailboxSettings, mySite, pastProjects, preferredName, responsibilities, schools, skills

$PropertiesToSelect = "accountEnabled,ageGroup,assignedLicenses,assignedPlans,businessPhones,city,companyName,consentProvidedForMinor,country,
createdDateTime,creationType,customSecurityAttributes,deletedDateTime,department,displayName,employeeHireDate,employeeLeaveDateTime,
employeeId,employeeOrgData,employeeType,externalUserState,externalUserStateChangeDateTime,faxNumber,givenName,id,identities,imAddresses,
jobTitle,lastPasswordChangeDateTime,legalAgeGroupClassification,licenseAssignmentStates,manager,mail,mailNickname,mobilePhone,officeLocation,
onPremisesDistinguishedName,onPremisesDomainName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,
onPremisesProvisioningErrors,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesUserPrincipalName,
otherMails,passwordPolicies,passwordProfile,postalCode,preferredDataLocation,preferredLanguage,provisionedPlans,proxyAddresses,
refreshTokensValidFromDateTime,serviceProvisioningErrors,securityIdentifier,showInAddressList,signInSessionsValidFromDateTime,
state,streetAddress,surname,usageLocation,userPrincipalName,userType"

# list available application extension properties
$AvailableExtensionProperties = Invoke-MgGraphAPI -Method "POST" -Endpoint "v1.0/directoryObjects/getAvailableExtensionProperties"
# filter available application extension properties for user objects
$AvailableUserExtensionProperties = $AvailableExtensionProperties.value | Where-Object {$_.targetObjects -contains "User"}
# add available application extension properties for user objects to the properties to select
if ($AvailableUserExtensionProperties.count -ge 1) {
    foreach ($AvailableUserExtensionProperty in $AvailableUserExtensionProperties) {
        $PropertiesToSelect = $PropertiesToSelect + ", $($AvailableUserExtensionProperty.name)"
    }
}

# list subscribed license SKUs
# https://learn.microsoft.com/en-us/graph/api/subscribedsku-list?view=graph-rest-1.0&tabs=http
Write-Host "[$(Get-Date -Format "HH:mm:ss")][INFO]: Getting all license info..."
$SubscribedSKUs = Invoke-MgGraphAPI -Endpoint "v1.0/subscribedSkus"
if ($SubscribedSKUs) {Write-Host -ForegroundColor Green "[$(Get-Date -Format "HH:mm:ss")][SUCCESS]:Successfully retrieved license info"}

# create hashtable of subscribed SKUs for efficient lookups
$SubscribedSKUsHashTable = @{}
foreach ($Entry in $SubscribedSKUs) {
    $SubscribedSKUsHashTable[$Entry.skuId] = $Entry
}

# if the tenant has at least one license plan including Entra ID Premium P1/P2, get signInActivity for the users as well
if ($SubscribedSKUs.servicePlans.servicePlanId -contains "eec0eb4f-6444-4f95-aba0-50c24d67f998") {
    # P2
    $PropertiesToSelect = $PropertiesToSelect + ",signInActivity"
} elseif ($SubscribedSKUs.servicePlans.servicePlanId -contains "41781fb2-bc02-4b7c-bd55-b576c07bb09d") {
    # P1
    $PropertiesToSelect = $PropertiesToSelect + ",signInActivity"
} else {
    # Free
}

# list users
# https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
# user resource: https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0
Write-Host "[$(Get-Date -Format "HH:mm:ss")][INFO]: Getting all users..."
$Users = Invoke-MgGraphAPI -Endpoint "v1.0/users?`$select=$($PropertiesToSelect)&`$top=999"
if ($Users) {Write-Host -ForegroundColor Green "[$(Get-Date -Format "HH:mm:ss")][SUCCESS]:Successfully retrieved $($Users.count) users"}

# if parameter 'IncludeRecipientType' or 'IncludeMailboxDetails' is provided
# -> grouped in the parameter set 'EXO'
# list recipients
if ($PSCmdlet.ParameterSetName -eq 'EXO') {

    Write-Host "[$(Get-Date -Format "HH:mm:ss")][INFO]: Getting all recipients..."

    $Recipients = Get-Recipient -ResultSize Unlimited -IncludeSoftDeletedRecipients

    # create hashtable of recipients for efficient lookups
    $RecipientsHashTable = @{}
    foreach ($Recipient in $Recipients) {
        $RecipientsHashTable[$Recipient.ExternalDirectoryObjectId] = $Recipient.RecipientTypeDetails
    }

    if ($Recipient) {Write-Host -ForegroundColor Green "[$(Get-Date -Format "HH:mm:ss")][SUCCESS]:Successfully retrieved $($Recipients.count) recipients"}

}

# if parameter 'IncludeMailboxDetails' is provided
if ($PSBoundParameters.ContainsKey('IncludeMailboxDetails')) {

    Write-Host "[$(Get-Date -Format "HH:mm:ss")][INFO]: Getting all mailbox details..."

    # https://learn.microsoft.com/en-us/graph/api/reportroot-getmailboxusagedetail?view=graph-rest-1.0&tabs=http
    # call the Graph API to get the URI where the report will be made available
    # -OutputType = HttpResponseMessage - returns the HTTP response without the content
    $MailboxUsageReportLocation = Invoke-MgGraphRequest -Method "GET" -Uri "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D180')" -OutputType HttpResponseMessage

    # use the URI returned by the Graph API to fetch the content of the report
    $MailboxUsageReportContent = Invoke-WebRequest -Uri $MailboxUsageReportLocation.RequestMessage.RequestUri.AbsoluteUri
    $MailboxUsageReport = [System.Text.Encoding]::UTF8.GetString($MailboxUsageReportContent.Content) | ConvertFrom-Csv

    # create hashtable for efficient lookups
    $MailboxDetailsHashTable = @{}
    foreach ($MailboxDetails in $MailboxUsageReport) {
        $MailboxDetailsHashTable[$MailboxDetails.'User Principal Name'] = $MailboxDetails
    }

    # https://learn.microsoft.com/en-us/graph/api/reportroot-getemailactivityuserdetail?view=graph-rest-1.0&tabs=http
    # call the Graph API to get the URI where the report will be made available
    # -OutputType = HttpResponseMessage - returns the HTTP response without the content
    $MailActivityReportLocation = Invoke-MgGraphRequest -Method "GET" -Uri "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D180')" -OutputType HttpResponseMessage

    # use the URI returned by the Graph API to fetch the content of the report
    $MailboxUsageReportContent = Invoke-WebRequest -Uri $MailboxUsageReportLocation.RequestMessage.RequestUri.AbsoluteUri
    $MailboxUsageReport = [System.Text.Encoding]::UTF8.GetString($MailboxUsageReportContent.Content) | ConvertFrom-Csv

}

# if parameter 'IncludeOneDriveDetails' is provided 
if ($PSBoundParameters.ContainsKey('IncludeOneDriveDetails')) {

    Write-Host "[$(Get-Date -Format "HH:mm:ss")][INFO]: Getting all OneDrive details..."

    # https://learn.microsoft.com/en-us/graph/api/reportroot-getonedriveusageaccountdetail?view=graph-rest-1.0&tabs=http
    # call the Graph API to get the URI where the report will be made available
    # -OutputType = HttpResponseMessage - returns the HTTP response without the content
    $OneDriveUsageReportLocation = Invoke-MgGraphRequest -Method "GET" -Uri "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D180')" -OutputType HttpResponseMessage

    # use the URI returned by the Graph API to fetch the content of the report
    $OneDriveUsageReportContent = Invoke-WebRequest -Uri $OneDriveUsageReportLocation.RequestMessage.RequestUri.AbsoluteUri
    $OneDriveUsageReport = [System.Text.Encoding]::UTF8.GetString($OneDriveUsageReportContent.Content) | ConvertFrom-Csv
    
    # create hashtable for efficient lookups
    $OneDriveDetailsHashTable = @{}
    foreach ($OneDriveDetails in $OneDriveUsageReport) {
        $OneDriveDetailsHashTable[$OneDriveDetails.'Owner Principal Name'] = $OneDriveDetails
    }

}

#region filtering 
if ($PSBoundParameters.ContainsKey('ImportCSVFilePath')) {
    $ImportCSVFile = Import-Csv -Path $ImportCSVFilePath -Delimiter ","
    $FilteredUsers = $Users | Where-Object {$_.id -in ($ImportCSVFile.ID) -or $_.userPrincipalName -in ($ImportCSVFile.ID)}
} else {
    $FilteredUsers = $Users
}

#endregion filtering

$i = 0
$Export = [System.Collections.Generic.List[Object]]@()

foreach ($User in $FilteredUsers) {

    $i++
    Write-Progress -Activity "Processing users" -PercentComplete (($i*100)/$FilteredUsers.count) -Status "$(([math]::Round((($i)/$FilteredUsers.count * 100),0))) %"

    $ExportObject = [PSCustomObject]@{
        Id                                          = $User.id
        UserPrincipalName                           = $User.userPrincipalName
        DisplayName                                 = $User.displayName
        GivenName                                   = $User.givenName
        SurName                                     = $User.surName
        AccountEnabled                              = $User.accountEnabled
        UserType                                    = $User.userType   
        OnPremisesSyncEnabled                       = if ($User.onPremisesSyncEnabled) {"TRUE"} else {"FALSE"} # only TRUE value is set, so blank = FALSE
        OnPremisesLastSyncDateTime                  = if ($User.onPremisesLastSyncDateTime) {$User.onPremisesLastSyncDateTime.ToString("yyyy-MM-dd HH:mm:ss")} else {"NA"}
        OnPremisesProvisioningErrors                = if ($User.onPremisesProvisioningErrors) {$User.onPremisesProvisioningErrors} else {"NA"}
        ProxyAddresses                              = $User.proxyAddresses -join ";"
        Mail                                        = $User.mail
        MailNickname                                = $User.mailNickname
        OtherMails                                  = $User.otherMails -join ";"
        IMAddresses                                 = $User.imAddresses -join ";"
        AssignedLicenses                            = "" # filled in later
        #AssignedPlans                              = $User.AssignedPlans # to review
        #ProvisionedPlans                           = $User.ProvisionedPlans # to review
        #LicenseAssignmentStates                    = $User.licenseAssignmentStates
        CompanyName                                 = $User.companyName
        CreatedDateTime                             = if ($User.createdDateTime) {$User.createdDateTime.ToString("yyyy-MM-dd HH:mm:ss")} else {"NA"}
        LastPasswordChangeDateTime                  = if ($User.LastPasswordChangeDateTime) {$User.LastPasswordChangeDateTime.ToString("yyyy-MM-dd HH:mm:ss")} else {"NA"}
        PasswordAgeInDays                           = if ($User.LastPasswordChangeDateTime) {([math]::Round((Get-Date - $User.LastPasswordChangeDateTime).TotalDays,0))} else {"NA"}
        SignInSessionsValidFromDateTime             = if ($User.signInSessionsValidFromDateTime) {$User.signInSessionsValidFromDateTime.ToString("yyyy-MM-dd HH:mm:ss")} else {"NA"}
        RefreshTokensValidFromDateTime              = if ($User.refreshTokensValidFromDateTime) {$User.refreshTokensValidFromDateTime.ToString("yyyy-MM-dd HH:mm:ss")} else {"NA"}
        LastSignInRequestId                         = "NA"
        LastSignInDateTime                          = "NA"
        LastNonInteractiveSignInRequestId           = "NA"
        LastNonInteractiveSignInDateTime            = "NA"
        CustomSecurityAttributes                    = $User.customSecurityAttributes
        OnPremisesUserPrincipalName                 = $User.onPremisesUserPrincipalName
        OnPremisesSamAccountName                    = $User.onPremisesSamAccountName
        OnPremisesDistinguishedName                 = $User.onPremisesDistinguishedName
        #OnPremisesExtensionAttributes               = $User.onPremisesExtensionAttributes
        OnPremisesImmutableId                       = $User.onPremisesImmutableId
        OnPremisesSecurityIdentifier                = $User.onPremisesSecurityIdentifier
        OnPremisesDomainName                        = $User.onPremisesDomainName
        SecurityIdentifier                          = $User.securityIdentifier
        RecipientTypeDetails                        = ""
        ShowInAddressList                           = $User.showInAddressList
        PasswordPolicies                            = $User.passwordPolicies
        PasswordProfile                             = if ($User.passwordProfile) {foreach ($Property in $User.passwordProfile.GetEnumerator()) {"$($Property.Key) = $($Property.value)`n"}} else {"NA"}
        JobTitle                                    = $User.jobTitle
        Department                                  = if ($User.streetAddress) {$User.streetAddress.ToString() -replace '[\r\n]',' '} else {""} # explicit conversion to string to deal with newline characters
        EmployeeId                                  = $User.employeeId
        EmployeeType                                = $User.employeeType
        EmployeeHireDate                            = $User.employeeHireDate
        EmployeeLeaveDateTime                       = $User.employeeLeaveDateTime
        EmployeeOrgData                             = $User.employeeOrgData
        MobilePhone                                 = $User.mobilePhone
        BusinessPhones                              = if ($User.businessPhones) {$User.businessPhones} else {"NA"}
        FaxNumber                                   = $User.faxNumber
        PreferredLanguage                           = $User.preferredLanguage
        PreferredDataLocation                       = $User.preferredDataLocation
        UsageLocation                               = $User.usageLocation
        CreationType                                = $User.creationType
        IdentitiesSignInType                        = if ($User.identities.signInType) {$User.identities.signInType -join ","} else {"NA"}
        IdentitiesIssuer                            = if ($User.identities.issuer) {$User.identities.issuer -join ","} else {"NA"}
        IdentitiesIssuerAssignedId                  = if ($User.identities.issuerAssignedId) {$User.identities.issuerAssignedId -join ","} else {"NA"}
        ExternalUserState                           = $User.externalUserState
        ExternalUserStateChangeDateTime             = $User.externalUserStateChangeDateTime
        OfficeLocation                              = $User.officeLocation
        StreetAddress                               = if ($User.streetAddress) {$User.streetAddress.ToString() -replace '[\r\n]',' '} else {""} # explicit conversion to string to deal with newline characters
        City                                        = $User.city
        PostalCode                                  = $User.postalCode
        State                                       = $User.state
        country                                     = $User.country
        AgeGroup                                    = $User.ageGroup
        LegalAgeGroupClassification                 = $User.legalAgeGroupClassification
        ConsentProvidedForMinor                     = $User.consentProvidedForMinor
        DeletedDateTime                             = if ($User.deletedDateTime) {$User.deletedDateTime.ToString("yyyy-MM-dd HH:mm:ss")} else {"NA"}
        #ServiceProvisioningErrors                   = $User.serviceProvisioningErrors
    }

    if ($AvailableUserExtensionProperties.count -ge 1) {
        foreach ($AvailableUserExtensionProperty in $AvailableUserExtensionProperties) {
            $ExportObject | Add-Member -NotePropertyName $($AvailableUserExtensionProperty.name) -NotePropertyValue $User.($($AvailableUserExtensionProperty.name))
        }
    }

    # signInActivity
    if ($subscribedSKUs.servicePlans.servicePlanName -contains "AAD_PREMIUM") {
        $ExportObject.LastSignInRequestId                   = if ($User.signInActivity.LastSignInRequestId) {$User.signInActivity.LastSignInRequestId} else {"none"}
        $ExportObject.LastSignInDateTime                    = if ($User.signInActivity.lastSignInDateTime) {$User.signInActivity.lastSignInDateTime.ToString("yyyy-MM-dd HH:mm:ss")} else {"none"}
        $ExportObject.LastNonInteractiveSignInRequestId     = if ($User.signInActivity.lastNonInteractiveSignInRequestId) {$User.signInActivity.lastNonInteractiveSignInRequestId} else {"none"}
        $ExportObject.LastNonInteractiveSignInDateTime      = if ($User.signInActivity.lastNonInteractiveSignInDateTime) {$User.signInActivity.lastNonInteractiveSignInDateTime.ToString("yyyy-MM-dd HH:mm:ss")} else {"none"}
    }

    # licenses
    if ($User.AssignedLicenses) {
        $AssignedLicensesList = [System.Collections.Generic.List[String]]@()
        foreach ($AssignedLicense in $User.AssignedLicenses) {
            $AssignedLicensesList.Add(($SubscribedSKUsHashTable[$AssignedLicense.skuId]).skuPartNumber)
        }
        $ExportObject.AssignedLicenses = ($AssignedLicensesList -join ";").ToString()
    }

    # App assignments
    if ($IncludeAppAssignments) {

        #
        # TODO
        #

        #$AppAssignments = Invoke-MgGraphAPI -Endpoint "v1.0/users/$($User.id)/AppRoleAssignments"
    }

    # group memberships
    if ($PSBoundParameters.ContainsKey('IncludeGroupMemberships')) {
        # add the property to the export object
        $exportObject | Add-Member -MemberType NoteProperty -Name MemberOf -Value "none"
        # get all direct group memberships of the user
        $directGroupMemberships = Invoke-MgGraphAPI -Endpoint "v1.0/users/$($user.id)/memberOf"
        # if user is member of at least 1 group
        if ($directGroupMemberships.count -ge 1) {
            $exportObject.MemberOf = '"' + (($directGroupMemberships | ForEach-Object { '{0} ({1})' -f $_.displayName, $_.Id }) -join ', ') + '"'
        }
    }

    # if parameter 'IncludeRecipientType' or 'IncludeMailboxDetails' is provided
    # -> grouped in the parameter set 'EXO'
    if ($PSCmdlet.ParameterSetName -eq 'EXO') {

        Write-Verbose "Getting recipient type for $($User.userPrincipalName)"

        if ($RecipientsHashTable[$($User.Id)]) {

            $ExportObject.RecipientTypeDetails = $RecipientsHashTable[$($User.Id)]

        } else {

            $ExportObject.RecipientTypeDetails = "NA"

        }
    }

    # if parameter 'IncludeMailboxDetails' is provided
    if ($PSBoundParameters.ContainsKey('IncludeMailboxDetails')) {

        $ExportObject | Add-Member -NotePropertyName MailboxStorageUsed         -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName MailboxArchive             -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName MailboxItemCount           -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName MailboxWarningQuota        -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName MailboxLastActivity        -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName MailboxDeletedItemSize     -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName MailboxDeletedItemCount    -NotePropertyValue "NA"

        Write-Verbose "Getting mailbox details for $($User.userPrincipalName)"

        if ($MailboxDetailsHashTable[$($User.userPrincipalName)]) {

            Write-Verbose "Mailbox details found for $($User.userPrincipalName)"

            $UserMailboxDetails = $MailboxDetailsHashTable[$($User.userPrincipalName)]
            
            <# sample
                Report Refresh Date                : 2024-09-13
                User Principal Name                : glenn@vanrymenant.eu
                Display Name                       : Glenn Van Rymenant
                Is Deleted                         : False
                Deleted Date                       : 
                Created Date                       : 2018-03-05
                Last Activity Date                 : 2024-09-13
                Item Count                         : 60481
                Storage Used (Byte)                : 35794878026
                Issue Warning Quota (Byte)         : 52613349376
                Prohibit Send Quota (Byte)         : 53150220288
                Prohibit Send/Receive Quota (Byte) : 53687091200
                Deleted Item Count                 : 1983
                Deleted Item Size (Byte)           : 14261431
                Deleted Item Quota (Byte)          : 32212254720
                Has Archive                        : False
                Report Period                      : 180
            #>

            $ExportObject.MailboxStorageUsed            = ([math]::round($UserMailboxDetails.'Storage Used (Byte)'/1Gb,2) -replace ",",".") + "GB"
            $ExportObject.MailboxArchive                = ($UserMailboxDetails.'Has Archive').ToString().ToUpper()
            $ExportObject.MailboxItemCount              = $UserMailboxDetails.'Item Count'
            $ExportObject.MailboxWarningQuota           =    ([math]::round($UserMailboxDetails.'Issue Warning Quota (Byte)'/1Gb,2) -replace ",",".") + "GB"
            $ExportObject.MailboxLastActivity           = $UserMailboxDetails.'Last Activity Date'
            $ExportObject.MailboxDeletedItemSize        = ([math]::round($UserMailboxDetails.'Deleted Item Size (Byte)'/1Gb,2) -replace ",",".") + "GB"
            $ExportObject.MailboxDeletedItemCount       = $UserMailboxDetails.'Deleted Item Count'
        }
    }

    # if parameter 'IncludeOneDriveDetails' is provided 
    if ($PSBoundParameters.ContainsKey('IncludeOneDriveDetails')) {

        $ExportObject | Add-Member -NotePropertyName OneDriveStorageUsed        -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName OneDriveFileCount          -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName OneDriveActiveFileCount    -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName OneDriveStorageAllocated   -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName OneDriveLastActivity       -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName OneDriveSiteId             -NotePropertyValue "NA"
        $ExportObject | Add-Member -NotePropertyName OneDriveSiteUrl            -NotePropertyValue "NA"

        Write-Verbose "Getting OneDrive details for $($User.userPrincipalName)"

        if ($OneDriveDetailsHashTable[$($User.userPrincipalName)]) {

            Write-Verbose "OneDrive details found for $($User.userPrincipalName)"

            $UserOneDriveDetails = $OneDriveDetailsHashTable[$($User.userPrincipalName)]
            
            <# sample
                Report Refresh Date      : 2024-09-13
                Site Id                  : a4be2312-cf80-4d8f-85b6-8615776ef256
                Site URL                 : ### doesn't seem to be returned yet - september 2024 ###
                Owner Display Name       : Glenn Van Rymenant
                Is Deleted               : False
                Last Activity Date       : 2024-09-11
                File Count               : 999
                Active File Count        : 123
                Storage Used (Byte)      : 313550758983
                Storage Allocated (Byte) : 1099511627776 (default 1024GB)
                Owner Principal Name     : glenn@vanrymenant.eu
                Report Period            : 180
            #>

            $ExportObject.OneDriveStorageUsed           = ([math]::round($UserOneDriveDetails.'Storage Used (Byte)'/1Gb,2) -replace ",",".") + "GB"
            $ExportObject.OneDriveFileCount             = $UserOneDriveDetails.'File Count'
            $ExportObject.OneDriveActiveFileCount       = $UserOneDriveDetails.'Active File Count'
            $ExportObject.OneDriveStorageAllocated      = ([math]::round($UserOneDriveDetails.'Storage Allocated (Byte)'/1Gb,2) -replace ",",".") + "GB"
            $ExportObject.OneDriveLastActivity          = $UserOneDriveDetails.'Last Activity Date'
            $ExportObject.OneDriveSiteId                = $UserOneDriveDetails.'Site Id'
            $ExportObject.OneDriveSiteURL               = $UserOneDriveDetails.'Site URL'  

        }
    }

    # add to export
    $Export.Add($ExportObject)  
    
}

# export to CSV
try {
    # ConvertTo-Csv > Export-Csv because ConvertTo-Csv will automatically quote all fields (columns) whereas Export-Csv will not
    $Export | ConvertTo-Csv -UseCulture | Out-File -Path $ExportPath -Encoding unicode
    Write-Host -ForegroundColor Green "[$(Get-Date -Format "HH:mm:ss")]: Successfully exported report to $($ExportPath)$($ExportFileName)" 
}
catch {
    throw "Failed to export CSV"
}

#endregion code

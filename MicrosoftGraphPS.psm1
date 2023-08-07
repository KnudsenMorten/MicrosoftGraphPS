Function Connect-MicrosoftGraphPS
{
<#
    .SYNOPSIS
    Connect to Microsoft Graph (requires PS-module Microsoft Graph minimum v2.x)

    .DESCRIPTION
    Connect to Microsoft Graph using Azure App & Secret
    Connect to Microsoft Graph using Azure App & Certificate Thumprint
    Connect to Microsoft Graph using interactive login and scope

    .AUTHOR
    Morten Knudsen, Microsoft MVP - https://mortenknudsen.net

    .LINK
    https://github.com/KnudsenMorten/MicrosoftGraphPS

    .PARAMETER AppId
    This is the Azure app id
        
    .PARAMETER AppSecret
    This is the secret of the Azure app

    .PARAMETER TenantId
    This is the Azure AD tenant id

    .PARAMETER CertificateThumbprint
    This is the thumprint of the installed certificate

    .PARAMETER ShowMgContext
    switch to show the current Microsoft Graph context

    .PARAMETER ShowMgContextExpandScopes
    switch to show the Microsoft Graph permissions in the current context

    .PARAMETER Scopes
    Here you can define an array of permissions

    .INPUTS
    None. You cannot pipe objects

    .OUTPUTS
    Connection to Microsoft Graph ("welcome")

    .EXAMPLE

    # Microsoft Graph connect with AzApp & Secret
    Connect-MicrosoftGraphPS -AppId $global:HighPriv_Modern_ApplicationID_Azure `
                             -AppSecret $global:HighPriv_Modern_Secret_Azure `
                             -TenantId $global:AzureTenantID

    # Microsoft Graph connect with AzApp & CertificateThumprint
    Connect-MicrosoftGraphPS -AppId $global:HighPriv_Modern_ApplicationID_Azure `
                             -CertificateThumbprint $global:HighPriv_Modern_CertificateThumbprint_Azure `
                             -TenantId $global:AzureTenantID

    # Show Permissions in the current context
    Connect-MicrosoftGraphPS -ShowMgContextExpandScopes

    # Show context of current Microsoft Graph context
    Connect-MicrosoftGraphPS -ShowMgContext

    # Microsoft Graph connect with interactive login with the permission defined in the scopes
    $Scopes = @("DeviceManagementConfiguration.ReadWrite.All",`
                "DeviceManagementManagedDevices.ReadWrite.All",`
                "DeviceManagementServiceConfig.ReadWrite.All"
                )
    Connect-MicrosoftGraphPS -Scopes $Scopes
#>

    [CmdletBinding()]
    param(
            [Parameter()]
                [string]$AppId,
            [Parameter()]
                [string]$AppSecret,
            [Parameter()]
                [string]$CertificateThumbprint,
            [Parameter()]
                [string]$TenantId,
            [Parameter()]
                [switch]$ShowMgContext = $false,
            [Parameter()]
                [switch]$ShowMgContextExpandScopes = $false,
            [Parameter()]
                [array]$Scopes
         )

    #---------------------------------------------------------------------
    # Microsoft Graph (MgGraph) connect with AzApp & AppSecret
    #---------------------------------------------------------------------

    If ( ($AppId) -and ($AppSecret) -and ($TenantId) )
        {
            $Disconnect = Disconnect-MgGraph -ErrorAction SilentlyContinue

            $AppSecretSecure = ConvertTo-SecureString $AppSecret -AsPlainText -Force
            $ClientSecretCredential = New-Object System.Management.Automation.PSCredential ($AppId, $AppSecretSecure)

            write-host "Connecting to Microsoft Graph using Azure App & Secret"
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential
        }

    #---------------------------------------------------------------------
    # Microsoft Graph (MgGraph) connect with AzApp & CertificateThumpprint
    #---------------------------------------------------------------------
    ElseIf ( ($AppId) -and ($CertificateThumbprint) -and ($TenantId) )
        {
            $Disconnect = Disconnect-MgGraph -ErrorAction SilentlyContinue

            write-host "Connecting to Microsoft Graph using Azure App & CertificateThumprint"
            Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $CertificateThumbprint
        }
    #---------------------------------------------------------------------
    # Microsoft Graph (MgGraph) connect using interactive connectivity
    #---------------------------------------------------------------------
    ElseIf ($ShowMgContext)
        {
            $Context = Get-MgContext
            Return $Context
        }

    ElseIf ($ShowMgContextExpandScopes)
        {
            $Context = Get-MgContext | Select -ExpandProperty Scopes
            Return $Context
        }
    Else
        {
            Connect-MgGraph -Scopes $Scopes
        }
}


Function Connect-MicrosoftRestApiEndpointPS
{
<#
.SYNOPSIS
Connect to REST API endpoint

.DESCRIPTION
Connect to REST API endpoint like https://api.securitycenter.microsoft.com

.AUTHOR
Morten Knudsen, Microsoft MVP - https://mortenknudsen.net

.LINK
https://github.com/KnudsenMorten/MicrosoftGraphPS

.PARAMETER Uri
This is the Uri for the REST endpoint in Microsoft Graph

.PARAMETER AppId
This is the Azure app id
        
.PARAMETER AppSecret
This is the secret of the Azure app

.PARAMETER TenantId
This is the Azure AD tenant id

.INPUTS
None. You cannot pipe objects

.OUTPUTS
Connection Header & Token

.EXAMPLE
$ConnectAuth = Connect-MicrosoftRestApiEndpointPS -AppId $global:HighPriv_Modern_ApplicationID_O365 `
                                                  -AppSecret $global:HighPriv_Modern_Secret_O365 `
                                                  -TenantId $global:AzureTenantID `
                                                  -Uri "https://api.securitycenter.microsoft.com"


#>
[CmdletBinding()]
param(
        [Parameter(mandatory)]
            [string]$Uri,
        [Parameter()]
            [string]$AppId,
        [Parameter()]
            [string]$AppSecret,
        [Parameter()]
            [string]$TenantId
        )

<#  TROUBLESHOOTING
    $AppId     = $global:HighPriv_Modern_ApplicationID_O365
    $AppSecret = $global:HighPriv_Modern_Secret_O365
    $TenantId  = $global:AzureTenantID
    $Uri       = "https://api.securitycenter.microsoft.com"
#>

    # Get Token
        $oAuthUri = "https://login.microsoftonline.com/$($TenantID)/oauth2/token"
        $authBody = [Ordered] @{
                                 resource = $Uri
                                 client_id = $AppId
                                 client_secret = $AppSecret
                                 grant_type = 'client_credentials'
                               }

        $AuthResponse = Invoke-RestMethod -Method Post -Uri $oAuthUri -Body $authBody -ErrorAction Stop

        $Token = $AuthResponse.access_token

    # Set the WebRequest headers
        $Headers = @{
                        'Content-Type' = 'application/json'
                        Accept = 'application/json'
                        Authorization = "Bearer $token"
                    }

    Return $Token, $Headers
}


Function Get-MgUser-AllProperties-AllUsers
{
<#

.SYNOPSIS
Performs a Get-MgUser for all users retrieving all properties (except for certain properties which cannot be returned within a user collection). 
Manager property is being expanded

.DESCRIPTION
Get all properties for all users
Expands manager information
Excludes certain properties which cannot be returned within a user collection in bulk retrieval (*)

(*)
https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http#optional-query-parameters

The following properties are only supported when retrieving a single user: aboutMe, birthday, hireDate, interests, mySite, pastProjects, preferredName, 
responsibilities, schools, skills, mailboxSettings, DeviceEnrollmentLimit, print, SignInActivity


.AUTHOR
Morten Knudsen, Microsoft MVP - https://mortenknudsen.net

.LINK
https://github.com/KnudsenMorten/MicrosoftGraphPS

.INPUTS
None. You cannot pipe objects

.OUTPUTS
Returns the data

.EXAMPLE

$Result = Get-MgUser-AllProperties-AllUsers
$Result | fl

$Result.ManagerProperties | fl

#>
    [CmdletBinding()]
    param(
            [Parameter()]
                [switch]$All,
            [Parameter()]
                [string]$UserId
         )

    # Building list of Properties from first user found in Entra ID (prior named Azure AD)
        $PropertiesRaw = Get-MgUser -Top 1 | Get-Member -MemberType Property | select -ExpandProperty Name

    <#
        $BulkPropertyExclude
        Certain properties cannot be returned within a user collection. 
        https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http#optional-query-parameters

        The following properties are only supported when retrieving a single user: aboutMe, birthday, hireDate, interests, mySite, pastProjects, preferredName, 
        responsibilities, schools, skills, mailboxSettings.

        The following properties are not supported in personal Microsoft accounts and will be null: aboutMe, birthday, interests, mySite, pastProjects, preferredName, 
        responsibilities, schools, skills, streetAddress.
    #>

        $PropertyExclude = @("MailboxSettings",`
                             "DeviceEnrollmentLimit",`
                             "SignInActivity",`
                             "Print",`
                             "AboutMe",`
                             "Birthday",`
                             "HireDate",`
                             "Interests",`
                             "MySite",`
                             "PastProjects",`
                             "PreferredName",`
                             "Responsibilities",`
                             "Schools",`
                             "Skills"
                            )

    # Removing special properties from bulk-retrieval
        $Properties = $PropertiesRaw | Where-Object { ($_ -notin $PropertyExclude) }

    # Building array of properties to expand
        $PropertiesExpand = @("Manager")

    # Getting all data about users
        Write-Host "Getting all properties from all users in Entra ID (prior named Azure AD) .... Please Wait !"
        $EntraID_Users_ALL = Get-MgUser -All -Property $Properties -ExpandProperty $PropertiesExpand | Select-Object $Properties | `
                             Select *,@{Name = 'ManagerDisplayName'; Expression = {$_.Manager.AdditionalProperties.displayName}}, `
                                      @{Name = 'ManagerMail'; Expression = {$_.Manager.AdditionalProperties.mail}},`
                                      @{Name = 'ManagerProperties'; Expression = {$_.Manager.AdditionalProperties}}
        
        Return $EntraID_Users_ALL
}


Function InstallUpdate-MicrosoftGraphPS
{
<#
    .SYNOPSIS
    Install and Update MicrosoftGraphPS module

    .DESCRIPTION
    Install latest version of MicrosoftGraphPS, if not found
    Updates to latest version of MicrosoftGraphPS, if switch (-AutoUpdate) is set

    .AUTHOR
    Morten Knudsen, Microsoft MVP - https://mortenknudsen.net

    .LINK
    https://github.com/KnudsenMorten/MicrosoftGraphPS

    .PARAMETER Scope
    Scope where MicrosoftGraphPS module will be installed - can be AllUsers or CurrentUser
        
    .PARAMETER AutoUpdate
    MicrosoftGraphPS module will be updated to latest version, if switch (-AutoUpdate) is set

    .INPUTS
    None. You cannot pipe objects

    .OUTPUTS
    Installation / Update status

    .EXAMPLE

    InstallUpdate-MicrosoftGraphPS -Scope AllUsers -AutoUpdate
#>

param(
      [parameter()]
          [ValidateSet("CurrentUser","AllUsers")]
          $Scope = "AllUsers",
      [parameter()]
          [switch]$AutoUpdate = $False
     )

    # check for MicrosoftGraphPS
        $ModuleCheck = Get-Module -Name MicrosoftGraphPS -ListAvailable -ErrorAction SilentlyContinue
        If (!($ModuleCheck))
            {
                # check for NuGet package provider
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                Write-Output ""
                Write-Output "Checking Powershell PackageProvider NuGet ... Please Wait !"
                    if (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) 
                        {
                            Write-Host "OK - PackageProvider NuGet is installed"
                        } 
                    else 
                        {
                            try
                                {
                                    Write-Host "Installing NuGet package provider .. Please Wait !"
                                    Install-PackageProvider -Name NuGet -Scope $Scope -Confirm:$false -Force
                                }
                            catch [Exception] {
                                $_.message 
                                exit
                            }
                        }

                Write-Output "Powershell module MicrosoftGraphPS was not found !"
                Write-Output "Installing latest version from PsGallery in scope $Scope .... Please Wait !"

                Install-module -Name MicrosoftGraphPS -Repository PSGallery -Force -Scope $Scope
                import-module -Name MicrosoftGraphPS -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
            }

        Elseif ($ModuleCheck)
            {
                # sort to get highest version, if more versions are installed
                $ModuleCheck = Sort-Object -Descending -Property Version -InputObject $ModuleCheck
                $ModuleCheck = $ModuleCheck[0]

                Write-Output "Checking latest version at PsGallery for MicrosoftGraphPS module"
                $online = Find-Module -Name MicrosoftGraphPS -Repository PSGallery

                #compare versions
                if ( ([version]$online.version) -gt ([version]$ModuleCheck.version) ) 
                    {
                        Write-Output "Newer version ($($online.version)) detected"

                        If ($AutoUpdate -eq $true)
                            {
                                Write-Output "Updating MicrosoftGraphPS module .... Please Wait !"
                                Update-module -Name MicrosoftGraphPS -Force
                                import-module -Name MicrosoftGraphPS -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
                            }
                    }
                else
                    {
                        # No new version detected ... continuing !
                        Write-Output "OK - Running latest version"

                        $UpdateAvailable = $False
                        import-module -Name MicrosoftGraphPS -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
                    }
            }
}


Function Invoke-MgGraphRequestPS
{
<#
    .SYNOPSIS
    Invoke command to get/put/post/patch/delete data using Microsoft Graph REST endpoint

    .DESCRIPTION
    Get data using Microsoft Graph REST endpoint in case there is no PS-cmdlet available

    .AUTHOR
    Morten Knudsen, Microsoft MVP - https://mortenknudsen.net

    .LINK
    https://github.com/KnudsenMorten/MicrosoftGraphPS

    .PARAMETER Uri
    This is the Uri for the REST endpoint in Microsoft Graph
        
    .PARAMETER Method
    This is the method to handle the data (GET, PUT, DELETE, POST, PATCH)

    .PARAMETER OutPutType
    This is the output type

    .INPUTS
    None. You cannot pipe objects

    .OUTPUTS
    Returns the data

    .EXAMPLE
    # Method #1 - REST Endpoint
    $Uri        = "https://graph.microsoft.com/v1.0/devicemanagement/managedDevices"
    $Devices    = Invoke-MgGraphRequestPS -Uri $Uri -Method GET -OutputType PSObject

    # Method #2 - MgGraph cmdlet (prefered method, if available)
    $Devices = Get-MgDeviceManagementManagedDevice
    $Devices

#>
    [CmdletBinding()]
    param(
            [Parameter(mandatory)]
                [string]$Uri,
            [Parameter(mandatory)]
                [ValidateSet("GET", "DELETE", "POST", "PUT", "PATCH", IgnoreCase = $false)] 
                $Method = "GET",
            [Parameter(mandatory)]
                [ValidateSet("PSObject", "JSON", "HashTable", "HttpResponseMessage", IgnoreCase = $false)] 
                $OutputType = "PSObject"
         )

    $Result       = @()
    $ResultsCount = 0

    $ResultsRaw   = Invoke-MGGraphRequest -Method $Method -Uri $Uri ?-OutputType $OutPutType

    $Result       += $ResultsRaw.value
    $ResultsCount += ($ResultsRaw.value | Measure-Object).count
    Write-host "[ $($ResultsCount) ]  Getting data from $($Uri) using MgGraph"

    if (!([string]::IsNullOrEmpty($ResultsRaw.'@odata.nextLink'))) 
        { 
            do 
                { 
                    Try
                        {
                            $ResultsRaw    = Invoke-MGGraphRequest -Method $Method -Uri $ResultsRaw.'@odata.nextLink' ?-OutputType $OutPutType
                        }
                    Catch
                        {
                            Write-host "Errors occured - waiting 3 sec and then retrying"
                            Sleep -Seconds 3
                        }

                    $ResultsCount += ($ResultsRaw.value | Measure-Object).count
                    $Result       += $ResultsRaw.value
                    Write-host "[ $($ResultsCount) ]  Getting more data from $($Uri) using MgGraph"
                } 
            while (!([string]::IsNullOrEmpty($ResultsRaw.'@odata.nextLink'))) 
        } 
    Return $Result
}


Function Invoke-MicrosoftRestApiRequestPS
{
<#
    .SYNOPSIS
    Invoke command to get/put/post/patch/delete data using Microsoft REST API endpoint

    .DESCRIPTION
    Get data using Microsoft REST API endpoint like GET https://api.securitycenter.microsoft.com/api/machines

    .AUTHOR
    Morten Knudsen, Microsoft MVP - https://mortenknudsen.net

    .LINK
    https://github.com/KnudsenMorten/MicrosoftGraphPS

    .PARAMETER Uri
    This is the Uri for the REST endpoint in Microsoft Graph
        
    .PARAMETER Method
    This is the method to handle the data (GET, PUT, DELETE, POST, PATCH)

    .PARAMETER Header
    This is the Header coming from Connect-MicrosoftRestApiEndpointPS

    .INPUTS
    None. You cannot pipe objects

    .OUTPUTS
    Returns the data

    .EXAMPLE
    $Result = Invoke-MicrosoftRestApiRequestPS -Uri "https://api.securitycenter.microsoft.com/api/machines" `
                                               -Method GET `
                                               -Headers $ConnectAuth[1]

    # Show Result
    $Result
#>
    [CmdletBinding()]
    param(
            [Parameter(mandatory)]
                [string]$Uri,
            [Parameter(mandatory)]
                [ValidateSet("GET", "DELETE", "POST", "PUT", "PATCH", IgnoreCase = $false)] 
                $Method = "GET",
            [Parameter(mandatory)]
                [Object]$Headers
         )

<#   TROUBLESHOOTING
    $Uri     = "https://api.securitycenter.microsoft.com/api/machines"
    $Method  = "GET"
    $Headers = $ConnectAuth[1]
#>

    $ResponseAllRecords = @()
    Do
        {
            Write-host ""

                try 
                    {
                        $ResponseRaw = Invoke-WebRequest -Method $Method -Uri $Uri -Headers $Headers
                        $ResponseAllRecords += $ResponseRaw.content
                        $ResponseRawJSON = ($ResponseRaw | ConvertFrom-Json)
                        $ResultsCount += ( ($ResponseRaw.content | ConvertFrom-Json).value | Measure-Object).count
                        Write-host "[ $($ResultsCount) ]  Getting data from $($Uri) using REST Api endpoint"


                        if ($ResponseRawJSON.'@odata.nextLink')
                            {
                                $Uri = $ResponseRawJSON.'@odata.nextLink'
                            } 
                        else 
                            {
                                $Uri = $null
                            }

                    }
                catch 
                    {
                        Write-host ""
                        Write-host "StatusCode: " $_.Exception.Response.StatusCode.value__
                        Write-host "StatusDescription:" $_.Exception.Response.StatusDescription
                        Write-host ""
  
                        if ($_.ErrorDetails.Message)
                            {
                                Write-host ""
                                Write-host "Inner Error: $_.ErrorDetails.Message"
                                Write-host ""
                            }
  
                        # check for a specific error so that we can retry the request otherwise, set the url to null so that we fall out of the loop
                        if ($_.Exception.Response.StatusCode.value__ -eq 403 )
                            {
                                # just ignore, leave the url the same to retry but pause first
                                if ($retryCount -ge $maxRetries)
                                    {
                                        # not going to retry again
                                        $Uri = $null
                                        Write-host 'Not going to retry...'
                                    }
                                else 
                                    {
                                        $retryCount += 1
                                        write-host ""
                                        Write-host "Retry attempt $retryCount after a $pauseDuration second pause..."
                                        Write-host ""
                                        Start-Sleep -Seconds $pauseDuration
                                    }
                            }
                            else
                                {
                                    # not going to retry -- set the url to null to fall back out of the while loop
                                    $Uri = $null
                                }
                    }

        }
    while (!([string]::IsNullOrEmpty($Uri)))

    $Result = ($ResponseAllRecords | ConvertFrom-Json).value

    Return $Result
}


Function Manage-Version-Microsoft.Graph
{
<#
.SYNOPSIS
Version management of Microsoft.Graph PS modules

.DESCRIPTION
Installing latest version of Microsoft.Graph, if not found
Shows older installed versions of Microsoft.Graph
Checks if newer version if available from PSGallery of Microsoft.Graph
Automatic clean-up old versions of Microsoft.Graph
Update to latest version from PSGallery of Microsoft.Graph
Remove all versions of Microsoft.Graph

.AUTHOR
Morten Knudsen, Microsoft MVP - https://mortenknudsen.net

.LINK
https://github.com/KnudsenMorten/MicrosoftGraphPS

.PARAMETER Scope
Scope where MicrosoftGraphPS module will be installed - can be AllUsers (default) or CurrentUser
        
.PARAMETER CleanupOldMicrosoftGraphVersions
[switch] Removes old versions, if any found

.PARAMETER RemoveAllMicrosoftGraphVersions
[switch] Removes all versions of Microsoft.Graph (complete re-install)

.PARAMETER InstallLatestMicrosoftGraph
[switch] Install latest version of Microsoft.Graph from PSGallery, if new version detected

.PARAMETER ShowVersionDetails
[switch] Show version details (detailed)

.INPUTS
None. You cannot pipe objects

.OUTPUTS
Returns the data

.EXAMPLE

# Show details of installed Microsoft.Graph
Manage-Version-Microsoft.Graph

# Show details of installed Microsoft.Graph including version details
Manage-Version-Microsoft.Graph -ShowVersionDetails

# Show details of installed Microsoft.Graph and install latest (if found)
Manage-Version-Microsoft.Graph -InstallLatestMicrosoftGraph

# Show details of installed Microsoft.Graph and install latest (if found)
Manage-Version-Microsoft.Graph -InstallLatestMicrosoftGraph -Scope CurrentUser

# Show details of installed Microsoft.Graph and clean-up old versions (if found)
Manage-Version-Microsoft.Graph -CleanupOldMicrosoftGraphVersions

# Show details of installed Microsoft.Graph and remove all versions (complete re-install)
Manage-Version-Microsoft.Graph -RemoveAllMicrosoftGraphVersions

# Show details, install latest (if found) and clean-up old versions (if found)
Manage-Version-Microsoft.Graph -InstallLatestMicrosoftGraph -CleanupOldMicrosoftGraphVersions
#>
    [CmdletBinding()]
    param(
            [parameter()]
                [ValidateSet("CurrentUser","AllUsers")]
                $Scope = "AllUsers",
            [Parameter()]
                [switch]$CleanupOldMicrosoftGraphVersions = $false,
            [Parameter()]
                [switch]$RemoveAllMicrosoftGraphVersions = $false,
            [Parameter()]
                [switch]$InstallLatestMicrosoftGraph = $False,
            [Parameter()]
                [switch]$ShowVersionDetails = $False
         )

    #-----------------------------------------------------------------------------------------

    # Remove all versions of Microsoft.Graph

    If ($RemoveAllMicrosoftGraphVersions)
        {
            Write-host ""
            Write-Host "Removing all versions of Microsoft.Graph main module ... Please Wait !"
            Remove-Module Microsoft.Graph -Force -ErrorAction SilentlyContinue
            Uninstall-Module Microsoft.Graph -AllVersions -Force -ErrorAction SilentlyContinue


            # Remove all dependency modules from memory + uninstall
            $Retry = 0
            Do
                {
                    $Retry = 1 + $Retry

                    $LoadedModules = Get-Module Microsoft.Graph.* -ListAvailable -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'Microsoft.Graph.Authentication' }
                    $LoadedModules = $LoadedModules | Sort-Object -Property Name

                    # Modules found
                    If ($LoadedModules)
                        {
                            ForEach ($Module in $LoadedModules)
                                {
                                    Write-Host "Removing dependency module $($Module.Name) (version: $($Module.Version)) ... Please Wait !"
                                    Remove-Module -Name $Module.Name -force -ErrorAction SilentlyContinue
                                    Uninstall-Module -Name $Module.Name -force -ErrorAction SilentlyContinue

                                    # Sometimes uninstall-module doesn't clean-up correctly. This will ensure complete deletion of leftovers !
                                    $ModulePath = (get-item $Module.Path -ErrorAction SilentlyContinue).DirectoryName
                                    if ( ($ModulePath) -and (Test-Path $ModulePath) )
                                        {
                                            $Result = takeown /F $ModulePath /A /R
                                            $Result = icacls $modulePath /reset
                                            $Result = icacls $modulePath /grant Administrators:'F' /inheritance:d /T
                                            $Result = Remove-Item -Path $ModulePath -Recurse -Force -Confirm:$false
                                        }
                                }


                            $LoadedModules = Get-Module -Name "Microsoft.Graph.Authentication" -ListAvailable -ErrorAction SilentlyContinue
                            ForEach ($Module in $LoadedModules)
                                {
                                    Write-Host "Removing dependency module $($Module.Name) (version: $($Module.Version)) ... Please Wait !"
                                    Remove-Module -Name $Module.Name -force -ErrorAction SilentlyContinue
                                    Uninstall-Module -Name $Module.Name -force -ErrorAction SilentlyContinue

                                    # Sometimes uninstall-module doesn't clean-up correctly. This will ensure complete deletion of leftovers !
                                    $ModulePath = (get-item $Module.Path -ErrorAction SilentlyContinue).DirectoryName
                                    if ( ($ModulePath) -and (Test-Path $ModulePath) )
                                        {
                                            $Result = takeown /F $ModulePath /A /R
                                            $Result = icacls $modulePath /reset
                                            $Result = icacls $modulePath /grant Administrators:'F' /inheritance:d /T
                                            $Result = Remove-Item -Path $ModulePath -Recurse -Force -Confirm:$false
                                        }
                                }
                        }

                    # Verifying if all modules have been removed
                    $InstalledModules = Get-Module Microsoft.Graph.* -ErrorAction SilentlyContinue
                }
            Until ( ($LoadedModules -eq $null) -or ($Retry -eq 5) )
        }

    #-----------------------------------------------------------------------------------------

    Write-host ""
    Write-Host "Checking if Microsoft.Graph is installed"
        $Installed = Get-module Microsoft.Graph.* -ListAvailable

        If ($Installed)
            {    
                Write-host ""
                Write-Host "OK - Microsoft.Graph was detected"
                $FoundGraph = $true
            }
        Else
            {
                Write-host ""
                Write-Host "Microsoft.Graph was not found on this computer. "
                $FoundGraph = $false
                $NewerVersionDetected = $false
            }

    #-----------------------------------------------------------------------------------------
        If ($FoundGraph -eq $true)
            {
                Write-host ""
                Write-Host "Checking if you are running latest version of Microsoft.Graph from PSGallery ... Please Wait !"

                $InstalledVersions = Get-module Microsoft.Graph* -ListAvailable | Where-Object { ($_.ModuleType -eq "Manifest") -or ($_.ModuleType -eq "Script") }
                $InstalledVersionsCount = ( ($InstalledVersions | Group-Object -Property Version) | Measure-Object).count
                $LatestVersion = $InstalledVersions | Sort-Object Version -Descending | Select-Object -First 1

                If ($ShowVersionDetails)
                    {
                        Write-host ""
                        Write-Host "Installed versions of Microsoft.Graph are [ $($InstalledVersionsCount) ]"
                        $InstalledVersions | Group-Object -Property Version
                    }

                # Checking latest version in PSGallery
                $online = Find-Module -Name Microsoft.Graph -Repository PSGallery

                #compare versions
                $NewerVersionDetected = $false   # default
                if ( ([version]$online.version) -gt ([version]$LatestVersion.version) ) 
                    {
                        $NewerVersionDetected = $true
                        Write-host ""
                        Write-host "Newer version ($($online.version)) of Microsoft.Graph was detected in PSGallery"
                    }
                Else
                    {
                        # No new version detected ... continuing !
                        Write-host ""
                        Write-host "OK - Running latest version ($($LatestVersion.version)) of Microsoft.Graph"
                        $NewerVersionDetected = $false
                    }

            #-----------------------------------------------------------------------------------------

                Write-host ""
                Write-Host "Checking if you have any older versions of Microsoft.Graph installed that may conflict and should be removed"

                $VersionsCleanup = $InstalledVersions | Where-Object { [version]$_.Version -lt [version]$Online.Version }
                $VersionsCleanupCount = ( ($VersionsCleanup | Group-Object -Property Version) | Measure-Object).count

                If ($VersionsCleanupCount -gt 0)
                    {
                        Write-Host ""
                        Write-Host "You have $VersionsCleanupCount older versions of Microsoft.Graph to remove"
                        Write-Host ""
                    }
                ElseIf ($VersionsCleanupCount -eq 0)
                     {
                        Write-Host ""
                        Write-Host "OK - You have no older versions of Microsoft.Graph installed"
                        Write-Host ""
                     }

                If ($ShowVersionDetails)
                    {
                        $VersionsCleanup | Select-Object Version
                    }
            }

    #-----------------------------------------------------------------------------------------

    # Update to latest version of Microsoft Graph - or install if missing
        # Update
        If ( ($InstallLatestMicrosoftGraph) -and ($NewerVersionDetected -eq $true) -and ($FoundGraph -eq $true) )
            {
                Write-host ""
                Write-host "Updating to latest version $($online.version) of Microsoft.Graph from PSGallery ... Please Wait !"
                Update-module Microsoft.Graph -Force
            }

        # Re-install
        ElseIf ( ($InstallLatestMicrosoftGraph) -and ($NewerVersionDetected -eq $false) -and ($FoundGraph -eq $true) )
            {
                # Checking latest version in PSGallery
                $online = Find-Module -Name Microsoft.Graph -Repository PSGallery

                Write-host ""
                Write-Host "Re-installing latest version ($($online.version)) of Microsoft.Graph from PS Gallery ... Please Wait !"
                Install-module Microsoft.Graph -Scope $Scope -Force
            }
        
        # New install
        ElseIf ( ($InstallLatestMicrosoftGraph) -and ($NewerVersionDetected -eq $false) -and ($FoundGraph -eq $false) )
            {
                # Checking latest version in PSGallery
                $online = Find-Module -Name Microsoft.Graph -Repository PSGallery

                Write-host ""
                Write-Host "Installing latest version ($($online.version)) of Microsoft.Graph from PS Gallery ... Please Wait !"
                Install-module Microsoft.Graph -Scope $Scope -Force
            }

    #-----------------------------------------------------------------------------------------

    # Remove older versions of Microsoft.Graph

    If ( ($CleanupOldMicrosoftGraphVersions) -and ($VersionsCleanupCount -gt 0) )
        {
            Write-host ""
            ForEach ($ModuleRemove in $VersionsCleanup)
                {
                    Write-Host "Removing older version $($ModuleRemove.Version) of $($ModuleRemove.Name) ... Please Wait !"

                    Uninstall-module -Name $ModuleRemove.Name -RequiredVersion $ModuleRemove.Version -Force -ErrorAction SilentlyContinue

                    # Removing left-overs if uninstall doesn't complete task
                    $ModulePath = (get-item $ModuleRemove.Path -ErrorAction SilentlyContinue).DirectoryName
                    if ( ($ModulePath) -and (Test-Path $ModulePath) )
                        {
                            $Result = takeown /F $ModulePath /A /R
                            $Result = icacls $modulePath /reset
                            $Result = icacls $modulePath /grant Administrators:'F' /inheritance:d /T
                            $Result = Remove-Item -Path $ModulePath -Recurse -Force -Confirm:$false
                        }
                }
        }
}



# SIG # Begin signature block
# MIIXHgYJKoZIhvcNAQcCoIIXDzCCFwsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDvwfLAis3qIm9T
# Be1/hCBgGoMQzjnPe3Y3F1ZR/4BvLaCCE1kwggVyMIIDWqADAgECAhB2U/6sdUZI
# k/Xl10pIOk74MA0GCSqGSIb3DQEBDAUAMFMxCzAJBgNVBAYTAkJFMRkwFwYDVQQK
# ExBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQDEyBHbG9iYWxTaWduIENvZGUgU2ln
# bmluZyBSb290IFI0NTAeFw0yMDAzMTgwMDAwMDBaFw00NTAzMTgwMDAwMDBaMFMx
# CzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQD
# EyBHbG9iYWxTaWduIENvZGUgU2lnbmluZyBSb290IFI0NTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBALYtxTDdeuirkD0DcrA6S5kWYbLl/6VnHTcc5X7s
# k4OqhPWjQ5uYRYq4Y1ddmwCIBCXp+GiSS4LYS8lKA/Oof2qPimEnvaFE0P31PyLC
# o0+RjbMFsiiCkV37WYgFC5cGwpj4LKczJO5QOkHM8KCwex1N0qhYOJbp3/kbkbuL
# ECzSx0Mdogl0oYCve+YzCgxZa4689Ktal3t/rlX7hPCA/oRM1+K6vcR1oW+9YRB0
# RLKYB+J0q/9o3GwmPukf5eAEh60w0wyNA3xVuBZwXCR4ICXrZ2eIq7pONJhrcBHe
# OMrUvqHAnOHfHgIB2DvhZ0OEts/8dLcvhKO/ugk3PWdssUVcGWGrQYP1rB3rdw1G
# R3POv72Vle2dK4gQ/vpY6KdX4bPPqFrpByWbEsSegHI9k9yMlN87ROYmgPzSwwPw
# jAzSRdYu54+YnuYE7kJuZ35CFnFi5wT5YMZkobacgSFOK8ZtaJSGxpl0c2cxepHy
# 1Ix5bnymu35Gb03FhRIrz5oiRAiohTfOB2FXBhcSJMDEMXOhmDVXR34QOkXZLaRR
# kJipoAc3xGUaqhxrFnf3p5fsPxkwmW8x++pAsufSxPrJ0PBQdnRZ+o1tFzK++Ol+
# A/Tnh3Wa1EqRLIUDEwIrQoDyiWo2z8hMoM6e+MuNrRan097VmxinxpI68YJj8S4O
# JGTfAgMBAAGjQjBAMA4GA1UdDwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBQfAL9GgAr8eDm3pbRD2VZQu86WOzANBgkqhkiG9w0BAQwFAAOCAgEA
# Xiu6dJc0RF92SChAhJPuAW7pobPWgCXme+S8CZE9D/x2rdfUMCC7j2DQkdYc8pzv
# eBorlDICwSSWUlIC0PPR/PKbOW6Z4R+OQ0F9mh5byV2ahPwm5ofzdHImraQb2T07
# alKgPAkeLx57szO0Rcf3rLGvk2Ctdq64shV464Nq6//bRqsk5e4C+pAfWcAvXda3
# XaRcELdyU/hBTsz6eBolSsr+hWJDYcO0N6qB0vTWOg+9jVl+MEfeK2vnIVAzX9Rn
# m9S4Z588J5kD/4VDjnMSyiDN6GHVsWbcF9Y5bQ/bzyM3oYKJThxrP9agzaoHnT5C
# JqrXDO76R78aUn7RdYHTyYpiF21PiKAhoCY+r23ZYjAf6Zgorm6N1Y5McmaTgI0q
# 41XHYGeQQlZcIlEPs9xOOe5N3dkdeBBUO27Ql28DtR6yI3PGErKaZND8lYUkqP/f
# obDckUCu3wkzq7ndkrfxzJF0O2nrZ5cbkL/nx6BvcbtXv7ePWu16QGoWzYCELS/h
# AtQklEOzFfwMKxv9cW/8y7x1Fzpeg9LJsy8b1ZyNf1T+fn7kVqOHp53hWVKUQY9t
# W76GlZr/GnbdQNJRSnC0HzNjI3c/7CceWeQIh+00gkoPP/6gHcH1Z3NFhnj0qinp
# J4fGGdvGExTDOUmHTaCX4GUT9Z13Vunas1jHOvLAzYIwggbmMIIEzqADAgECAhB3
# vQ4DobcI+FSrBnIQ2QRHMA0GCSqGSIb3DQEBCwUAMFMxCzAJBgNVBAYTAkJFMRkw
# FwYDVQQKExBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQDEyBHbG9iYWxTaWduIENv
# ZGUgU2lnbmluZyBSb290IFI0NTAeFw0yMDA3MjgwMDAwMDBaFw0zMDA3MjgwMDAw
# MDBaMFkxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMS8w
# LQYDVQQDEyZHbG9iYWxTaWduIEdDQyBSNDUgQ29kZVNpZ25pbmcgQ0EgMjAyMDCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBANZCTfnjT8Yj9GwdgaYw90g9
# z9DljeUgIpYHRDVdBs8PHXBg5iZU+lMjYAKoXwIC947Jbj2peAW9jvVPGSSZfM8R
# Fpsfe2vSo3toZXer2LEsP9NyBjJcW6xQZywlTVYGNvzBYkx9fYYWlZpdVLpQ0LB/
# okQZ6dZubD4Twp8R1F80W1FoMWMK+FvQ3rpZXzGviWg4QD4I6FNnTmO2IY7v3Y2F
# QVWeHLw33JWgxHGnHxulSW4KIFl+iaNYFZcAJWnf3sJqUGVOU/troZ8YHooOX1Re
# veBbz/IMBNLeCKEQJvey83ouwo6WwT/Opdr0WSiMN2WhMZYLjqR2dxVJhGaCJedD
# CndSsZlRQv+hst2c0twY2cGGqUAdQZdihryo/6LHYxcG/WZ6NpQBIIl4H5D0e6lS
# TmpPVAYqgK+ex1BC+mUK4wH0sW6sDqjjgRmoOMieAyiGpHSnR5V+cloqexVqHMRp
# 5rC+QBmZy9J9VU4inBDgoVvDsy56i8Te8UsfjCh5MEV/bBO2PSz/LUqKKuwoDy3K
# 1JyYikptWjYsL9+6y+JBSgh3GIitNWGUEvOkcuvuNp6nUSeRPPeiGsz8h+WX4VGH
# aekizIPAtw9FbAfhQ0/UjErOz2OxtaQQevkNDCiwazT+IWgnb+z4+iaEW3VCzYkm
# eVmda6tjcWKQJQ0IIPH/AgMBAAGjggGuMIIBqjAOBgNVHQ8BAf8EBAMCAYYwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU
# 2rONwCSQo2t30wygWd0hZ2R2C3gwHwYDVR0jBBgwFoAUHwC/RoAK/Hg5t6W0Q9lW
# ULvOljswgZMGCCsGAQUFBwEBBIGGMIGDMDkGCCsGAQUFBzABhi1odHRwOi8vb2Nz
# cC5nbG9iYWxzaWduLmNvbS9jb2Rlc2lnbmluZ3Jvb3RyNDUwRgYIKwYBBQUHMAKG
# Omh0dHA6Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0L2NvZGVzaWduaW5n
# cm9vdHI0NS5jcnQwQQYDVR0fBDowODA2oDSgMoYwaHR0cDovL2NybC5nbG9iYWxz
# aWduLmNvbS9jb2Rlc2lnbmluZ3Jvb3RyNDUuY3JsMFYGA1UdIARPME0wQQYJKwYB
# BAGgMgEyMDQwMgYIKwYBBQUHAgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29t
# L3JlcG9zaXRvcnkvMAgGBmeBDAEEATANBgkqhkiG9w0BAQsFAAOCAgEACIhyJsav
# +qxfBsCqjJDa0LLAopf/bhMyFlT9PvQwEZ+PmPmbUt3yohbu2XiVppp8YbgEtfjr
# y/RhETP2ZSW3EUKL2Glux/+VtIFDqX6uv4LWTcwRo4NxahBeGQWn52x/VvSoXMNO
# Ca1Za7j5fqUuuPzeDsKg+7AE1BMbxyepuaotMTvPRkyd60zsvC6c8YejfzhpX0FA
# Z/ZTfepB7449+6nUEThG3zzr9s0ivRPN8OHm5TOgvjzkeNUbzCDyMHOwIhz2hNab
# XAAC4ShSS/8SS0Dq7rAaBgaehObn8NuERvtz2StCtslXNMcWwKbrIbmqDvf+28rr
# vBfLuGfr4z5P26mUhmRVyQkKwNkEcUoRS1pkw7x4eK1MRyZlB5nVzTZgoTNTs/Z7
# KtWJQDxxpav4mVn945uSS90FvQsMeAYrz1PYvRKaWyeGhT+RvuB4gHNU36cdZytq
# tq5NiYAkCFJwUPMB/0SuL5rg4UkI4eFb1zjRngqKnZQnm8qjudviNmrjb7lYYuA2
# eDYB+sGniXomU6Ncu9Ky64rLYwgv/h7zViniNZvY/+mlvW1LWSyJLC9Su7UpkNpD
# R7xy3bzZv4DB3LCrtEsdWDY3ZOub4YUXmimi/eYI0pL/oPh84emn0TCOXyZQK8ei
# 4pd3iu/YTT4m65lAYPM8Zwy2CHIpNVOBNNwwggb1MIIE3aADAgECAgx5Y9ljauM7
# cdkFAm4wDQYJKoZIhvcNAQELBQAwWTELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEds
# b2JhbFNpZ24gbnYtc2ExLzAtBgNVBAMTJkdsb2JhbFNpZ24gR0NDIFI0NSBDb2Rl
# U2lnbmluZyBDQSAyMDIwMB4XDTIzMDMyNzEwMjEzNFoXDTI2MDMyMzE2MTgxOFow
# YzELMAkGA1UEBhMCREsxEDAOBgNVBAcTB0tvbGRpbmcxEDAOBgNVBAoTBzJsaW5r
# SVQxEDAOBgNVBAMTBzJsaW5rSVQxHjAcBgkqhkiG9w0BCQEWD21va0AybGlua2l0
# Lm5ldDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAMykjWtM6hY5IRPe
# VIVB+yX+3zcMJQR2gjTZ81LnGVRE94Zk2GLFAwquGYWt1shoTHTV5j6Ef2AXYBDV
# kNruisJVJ17UsMGdsU8upwdZblFbLNzLw+qBXVC/OUVua9M0cub7CfUNkn/Won4D
# 7i41QyuDXdZFOIfRhZ3qnCYCJCSgYLoUXAS6xei2tPkkk1w8aXEFxybyy7eRqQjk
# HqIS5N4qH3YQkz+SbSlz/yj6mD65H5/Ts+lZxX2xL/8lgJItpdaJx+tarprv/tT+
# +n9a13P53YNzCWOmyhd376+7DMXxxSzT24kq13Ks3xnUPGoWUx2UPRnJHjTWoBfg
# Y7Zd3MffrdO0QEoDC9X5F5boh6oankVSOdSPRFns085KI+vkbt3bdG62MIeUbNtS
# v7mZBX8gcYv0szlo0ey7bbOJWoiZFT2fB+pBVvxDhpYP0/3aFveM1wfhshaJBhxx
# /2GCswYYBHH7B3+8j4BT8N8S030q4snys2Qt9tdFIHvSV7lIw/yorT1WM1cr+Lqo
# 74eR+Hi982db0k68p2BGdCOY0QhhaNqxufwbK+gVWrQY57GIX/1cUrBt0akMsli2
# 19xVmUGhIw85ZF7wcQplhslbUxyNUilY+c93q1bsIFjaOnjjvo56g+kyKICm5zsG
# FQLRVaXUSLY+i8NSiH8fd64etaptAgMBAAGjggGxMIIBrTAOBgNVHQ8BAf8EBAMC
# B4AwgZsGCCsGAQUFBwEBBIGOMIGLMEoGCCsGAQUFBzAChj5odHRwOi8vc2VjdXJl
# Lmdsb2JhbHNpZ24uY29tL2NhY2VydC9nc2djY3I0NWNvZGVzaWduY2EyMDIwLmNy
# dDA9BggrBgEFBQcwAYYxaHR0cDovL29jc3AuZ2xvYmFsc2lnbi5jb20vZ3NnY2Ny
# NDVjb2Rlc2lnbmNhMjAyMDBWBgNVHSAETzBNMEEGCSsGAQQBoDIBMjA0MDIGCCsG
# AQUFBwIBFiZodHRwczovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5LzAI
# BgZngQwBBAEwCQYDVR0TBAIwADBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3Js
# Lmdsb2JhbHNpZ24uY29tL2dzZ2NjcjQ1Y29kZXNpZ25jYTIwMjAuY3JsMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB8GA1UdIwQYMBaAFNqzjcAkkKNrd9MMoFndIWdkdgt4
# MB0GA1UdDgQWBBQxxpY2q5yrKa7VFODTZhTfPKmyyTANBgkqhkiG9w0BAQsFAAOC
# AgEAe38NgZR4IV9u264/n/jiWlHbBu847j1vpN6dovxMvdUQZ780eH3JzcvG8fo9
# 1uO1iDIZksSigiB+d8Sj5Yvh+oXlfYEffjIQCwcIlWNciOzWYZzl9qPHXgdTnaIu
# JA5cR846TepQLVMXc1Yb72Z7OGjldmRIxGjRimDsmzY+TdTu15lF4IkUj0VJhr8F
# PYOdEVZVOXHtPmUjPqsq9M7WpALYbc0pUawcy0FOOwXqzaCk7O3vMXej4Oycm6RB
# GfRH3JPOCvH2ddiIfPq2Lce4nhTuLsgumBJE2vOalVddIfTBjE9PpMub15lHyp1m
# fW0ZJvXOghPvRqufMT3SjPTHt6PV8LwhQD8BiGSZ9rp94js4xTnGexSOFKLLMxWE
# PTr5EPe3kmtspGgKCqLEZvsMYz7JlWNuaHBy+vdQZWV3376luwV4IHfGT+1wxe0E
# 90dMRI+9SNIKkVvKV3FUtToZUh3Np4cCIHJLQ1eslXFzIJa6wrjVsnWM/3OyedpQ
# JERGNYXlVmxdgGFjrY1I6UWII0Y1iZW3t+JvhXosUaha8i/YSxaDH+5H/Klad2OZ
# Xq4Eg39QxkCELbmJmSU0sUYNnl0JTEu6jJY9UJMFikzf5s3p2ZuKdyMbRgN5GNNV
# 883meI/X5KVHBJDG1epigMer7fFXMVZUGoI12iIz/gOolQExggMbMIIDFwIBATBp
# MFkxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMS8wLQYD
# VQQDEyZHbG9iYWxTaWduIEdDQyBSNDUgQ29kZVNpZ25pbmcgQ0EgMjAyMAIMeWPZ
# Y2rjO3HZBQJuMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKA
# AKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIJTvkpt/HL/TRLe3nGqqSnMc
# uit3URpdCZlotZGCq1mMMA0GCSqGSIb3DQEBAQUABIICALfX7PXGTXZt7Un/CAnM
# NCiTe0mFpVXI3jYlbZ8cMVwBZN7PiI3QBn43Nn8zXcvmTGpVlqz31PA02Krpdr0Z
# ebprbH/ip4h9p8RTqzPHC41Ra3+8tCoWe9fd8qU3vqyyEzlTXVgFXccsazvd2pZb
# huTggmKe6N889Vh6Gb1tWOSwrcc9jWvDJI+qpqSYYchgbl4hrtEBKBMZeoQbySNc
# iYkriCWCu2dah4K54PjpuwQTbCET55ury8jwXG4dS5Ik/KOzZUtiPqiEX0NSUYlM
# lnSr2usqIHsIKi/R+E3cgp8DggLyPGPoicDD/BGBPCMssumVkJdjO/aXxUuIZH73
# gTajWovdTcpXZ5cgSEkgI93zqsFAVyjf08A5WnraRghp0uCr38fTte4OItSwRYiy
# 9qwp0BPxuIoGSYWwKYwLmb1rtMp+0IFl9mEFkHRyKTgDr9EdkS97QYN75a0BTcjv
# E8O7ZR9xD4emttqv7oaLUTiVn9bp2HAK6VaZCw5J7Rs7fi2gZ8vJSpGEuXqQVfis
# +qTcqAowlxm12uYLxM+wrJS+1vrblyuDQ4CZd9NFHRS+Wb4EVRvoR4MveK705srH
# btRvv8Q+EeygUWycQo4wn0IWebWrHnrW6af99EBmO1BRUSyeZl41/uo1sZ0yUKTQ
# aWOJ+U5vpLKKPe3ea1Ti2q9f
# SIG # End signature block

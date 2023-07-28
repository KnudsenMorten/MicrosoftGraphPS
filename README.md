# MicrosoftGraphPS
Think of this PS-module as a helper for **Microsoft Graph version-management, connectivity** and **data management** using **Microsoft Graph**. More functions will be added when needed.

| Function                       | Description                                                  |
| ------------------------------ | ------------------------------------------------------------ |
| Manage-Version-Microsoft.Graph | Version management of Microsoft.Graph PS modules<br/>Installing latest version of Microsoft.Graph, if not found<br/>Shows older installed versions of Microsoft.Graph<br/>Checks if newer version if available from PSGallery of Microsoft.Graph<br/>Automatic clean-up old versions of Microsoft.Graph<br/>Update to latest version from PSGallery of Microsoft.Graph |
| Connect-MicrosoftGraphPS       | Connect to Microsoft Graph using Azure App & Secret<br/>Connect to Microsoft Graph using Azure App & Certificate Thumprint<br/>Connect to Microsoft Graph using interactive login and scope |
| Invoke-MgGraphRequestPS        | Invoke command with pagination support to get/put/post/patch/delete data using Microsoft Graph REST endpoint. |
| InstallUpdate-MicrosoftGraphPS | Install latest version of MicrosoftGraphPS, if not found<br/>Update to latest version of MicrosoftGraphPS, if switch (-AutoUpdate) is set |



## Pre-req script for getting environment ready with Microsoft.Graph and MicrosoftGraphPS

Just copy the entire script-code below into the beginning of your script - and change the variables according to your needs as outlined below.

You can also [download the script here](https://raw.githubusercontent.com/KnudsenMorten/MicrosoftGraphPS/main/Install-Update-MicrosoftGraphPS-Microsoft.Graph.ps1). 

You can run the pre-req code as part of your script and it will be able to update to latest version and remove old versions, if desired.



##### Variables

```
$Scope      = "AllUsers"  # Valid parameters: AllUsers, CurrentUser
$AutoUpdate = $True
```

$Scope controls where MicrosoftGraphPS PS-module is installed (AllUsers, CurrentUser)

You can auto-update to latest version of MicrosoftGraphPS, if you set $AutoUpdate to $True. 

If you want to control which version, you can disable AutoUpdate ($AutoUpdate = $False)



**Complete script**

    ##########################################################################################
    # Pre-req script for getting environment ready with Microsoft.Graph and MicrosoftGraphPS
    ##########################################################################################
    
    <#
    .SYNOPSIS
    Install and Update MicrosoftGraphPS module
    Version management of Microsoft.Graph PS modules
    
    .DESCRIPTION
    
    MicrosoftGraphPS:
     Install latest version of MicrosoftGraphPS, if not found
     Updates to latest version of MicrosoftGraphPS, if switch ($AutoUpdate) is set to $True
    
    Microsoft.Graph:
     Installing latest version of Microsoft.Graph, if not found
     Shows older installed versions of Microsoft.Graph
     Checks if newer version if available from PSGallery of Microsoft.Graph
     Automatic clean-up old versions of Microsoft.Graph
     Update to latest version from PSGallery of Microsoft.Graph
    
    .AUTHOR
    Morten Knudsen, Microsoft MVP - https://mortenknudsen.net
    
    .LINK
    https://github.com/KnudsenMorten/MicrosoftGraphPS
    #>
    
    # Variables
    $Scope      = "AllUsers"  # Valid parameters: AllUsers, CurrentUser
    $AutoUpdate = $True
    
    # Check if MicrosoftGraphPS is installed
    $ModuleCheck = Get-Module -Name MicrosoftGraphPS -ListAvailable -ErrorAction SilentlyContinue
    
    If (!($ModuleCheck))    # MicrosoftGraphPS is NOT installed
        {
            # check for NuGet package provider
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
            Write-host ""
            Write-host "Checking Powershell PackageProvider NuGet ... Please Wait !"
                if (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) 
                    {
                        Write-host ""
                        Write-Host "OK - PackageProvider NuGet is installed"
                    } 
                else 
                    {
                        try
                            {
                                Write-host ""
                                Write-Host "Installing NuGet package provider .. Please Wait !"
                                Install-PackageProvider -Name NuGet -Scope $Scope -Confirm:$false -Force
                            }
                        catch [Exception] {
                            $_.message 
                            exit
                        }
                    }
    
            Write-host "Powershell module MicrosoftGraphPS was not found !"
            Write-Host ""
            Write-host "Installing latest version from PsGallery in scope $Scope .... Please Wait !"
            Write-Host ""
    
            Install-module -Name MicrosoftGraphPS -Repository PSGallery -Force -Scope $Scope
            import-module -Name MicrosoftGraphPS -Global -force -DisableNameChecking -WarningAction SilentlyContinue
        }
            
    Elseif ($ModuleCheck)    # MicrosoftGraphPS is installed - checking version, if it should be updated
        {
            # sort to get highest version, if more versions are installed
            $ModuleCheck = Sort-Object -Descending -Property Version -InputObject $ModuleCheck
            $ModuleCheck = $ModuleCheck[0]
    
            Write-host "Checking latest version of MicrosoftGraphPS module at PsGallery"
            $online = Find-Module -Name MicrosoftGraphPS -Repository PSGallery
    
            #compare versions
            if ( ([version]$online.version) -gt ([version]$ModuleCheck.version) ) 
                {
                    write-host ""
                    Write-host "   Newer version ($($online.version)) detected"
    
                    If ($AutoUpdate -eq $true)
                        {
                            write-host ""
                            Write-host "   Updating MicrosoftGraphPS module .... Please Wait !"
                            Write-Host ""
    
                            Update-module -Name MicrosoftGraphPS -Force
                            import-module -Name MicrosoftGraphPS -Global -force -DisableNameChecking -WarningAction SilentlyContinue
                        }
                }
            else
                {
                    # No new version detected ... continuing !
                    write-host ""
                    Write-host "   OK - Running latest version of MicrosoftGraphPS"
                    Write-Host ""
    
                    $UpdateAvailable = $False
                    import-module -Name MicrosoftGraphPS -Global -force -DisableNameChecking -WarningAction SilentlyContinue
                }
        }
    
    ##########################################################################################
    # Install-Update-Cleanup-Microsoft.Graph
    ##########################################################################################
    If ($AutoUpdate)
        {
            Manage-Version-Microsoft.Graph -InstallLatestMicrosoftGraph -CleanupOldMicrosoftGraphVersions -Scope $Scope
        }
    Else
        {
            Manage-Version-Microsoft.Graph -Scope $Scope
        }



# **Synopsis for functions**



## Manage-Version-Microsoft.Graph

```
.SYNOPSIS
Version management of Microsoft.Graph PS modules

.DESCRIPTION
Installing latest version of Microsoft.Graph, if not found
Shows older installed versions of Microsoft.Graph
Checks if newer version if available from PSGallery of Microsoft.Graph
Automatic clean-up old versions of Microsoft.Graph
Update to latest version from PSGallery of Microsoft.Graph

.AUTHOR
Morten Knudsen, Microsoft MVP - https://mortenknudsen.net

.LINK
https://github.com/KnudsenMorten/MicrosoftGraphPS

.PARAMETER Scope
Scope where MicrosoftGraphPS module will be installed - can be AllUsers or CurrentUser
        
.PARAMETER CleanupOldMicrosoftGraphVersions
[switch] Removes old versions, if any found

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

# Show details of installed Microsoft.Graph and clean-up old versions (if found)
Manage-Version-Microsoft.Graph -CleanupOldMicrosoftGraphVersions

# Show details, install latest (if found) and clean-up old versions (if found)
Manage-Version-Microsoft.Graph -InstallLatestMicrosoftGraph -CleanupOldMicrosoftGraphVersions

```



## Connect-MicrosoftGraphPS

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


## Invoke-MgGraphRequestPS

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



## InstallUpdate-MicrosoftGraphPS

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




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
        InstallUpdate-MicrosoftGraphPS -Scope AllUsers -AutoUpdate
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

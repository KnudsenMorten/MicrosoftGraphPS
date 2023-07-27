###################################
# Install-Update-MicrosoftGraphPS
###################################
<#
    .SYNOPSIS
    Install and Update MicrosoftGraphPS module

    .DESCRIPTION
    Install latest version of MicrosoftGraphPS, if not found
    Updates to latest version of MicrosoftGraphPS, if switch ($AutoUpdate) is set to $True

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
        
    Elseif ($ModuleCheck)    # MicrosoftGraphPS is installed - checking version, if it should be updated
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

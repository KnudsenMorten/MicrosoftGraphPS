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
# MIIRgwYJKoZIhvcNAQcCoIIRdDCCEXACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUVTpVKIom1z5kENBMm/OycsYj
# TLmggg3jMIIG5jCCBM6gAwIBAgIQd70OA6G3CPhUqwZyENkERzANBgkqhkiG9w0B
# AQsFADBTMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTEp
# MCcGA1UEAxMgR2xvYmFsU2lnbiBDb2RlIFNpZ25pbmcgUm9vdCBSNDUwHhcNMjAw
# NzI4MDAwMDAwWhcNMzAwNzI4MDAwMDAwWjBZMQswCQYDVQQGEwJCRTEZMBcGA1UE
# ChMQR2xvYmFsU2lnbiBudi1zYTEvMC0GA1UEAxMmR2xvYmFsU2lnbiBHQ0MgUjQ1
# IENvZGVTaWduaW5nIENBIDIwMjAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIK
# AoICAQDWQk3540/GI/RsHYGmMPdIPc/Q5Y3lICKWB0Q1XQbPDx1wYOYmVPpTI2AC
# qF8CAveOyW49qXgFvY71TxkkmXzPERabH3tr0qN7aGV3q9ixLD/TcgYyXFusUGcs
# JU1WBjb8wWJMfX2GFpWaXVS6UNCwf6JEGenWbmw+E8KfEdRfNFtRaDFjCvhb0N66
# WV8xr4loOEA+COhTZ05jtiGO792NhUFVnhy8N9yVoMRxpx8bpUluCiBZfomjWBWX
# ACVp397CalBlTlP7a6GfGB6KDl9UXr3gW8/yDATS3gihECb3svN6LsKOlsE/zqXa
# 9FkojDdloTGWC46kdncVSYRmgiXnQwp3UrGZUUL/obLdnNLcGNnBhqlAHUGXYoa8
# qP+ix2MXBv1mejaUASCJeB+Q9HupUk5qT1QGKoCvnsdQQvplCuMB9LFurA6o44EZ
# qDjIngMohqR0p0eVfnJaKnsVahzEaeawvkAZmcvSfVVOIpwQ4KFbw7MueovE3vFL
# H4woeTBFf2wTtj0s/y1KiirsKA8tytScmIpKbVo2LC/fusviQUoIdxiIrTVhlBLz
# pHLr7jaep1EnkTz3ohrM/Ifll+FRh2npIsyDwLcPRWwH4UNP1IxKzs9jsbWkEHr5
# DQwosGs0/iFoJ2/s+PomhFt1Qs2JJnlZnWurY3FikCUNCCDx/wIDAQABo4IBrjCC
# AaowDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMBIGA1UdEwEB
# /wQIMAYBAf8CAQAwHQYDVR0OBBYEFNqzjcAkkKNrd9MMoFndIWdkdgt4MB8GA1Ud
# IwQYMBaAFB8Av0aACvx4ObeltEPZVlC7zpY7MIGTBggrBgEFBQcBAQSBhjCBgzA5
# BggrBgEFBQcwAYYtaHR0cDovL29jc3AuZ2xvYmFsc2lnbi5jb20vY29kZXNpZ25p
# bmdyb290cjQ1MEYGCCsGAQUFBzAChjpodHRwOi8vc2VjdXJlLmdsb2JhbHNpZ24u
# Y29tL2NhY2VydC9jb2Rlc2lnbmluZ3Jvb3RyNDUuY3J0MEEGA1UdHwQ6MDgwNqA0
# oDKGMGh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5jb20vY29kZXNpZ25pbmdyb290cjQ1
# LmNybDBWBgNVHSAETzBNMEEGCSsGAQQBoDIBMjA0MDIGCCsGAQUFBwIBFiZodHRw
# czovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5LzAIBgZngQwBBAEwDQYJ
# KoZIhvcNAQELBQADggIBAAiIcibGr/qsXwbAqoyQ2tCywKKX/24TMhZU/T70MBGf
# j5j5m1Ld8qIW7tl4laaafGG4BLX468v0YREz9mUltxFCi9hpbsf/lbSBQ6l+rr+C
# 1k3MEaODcWoQXhkFp+dsf1b0qFzDTgmtWWu4+X6lLrj83g7CoPuwBNQTG8cnqbmq
# LTE7z0ZMnetM7LwunPGHo384aV9BQGf2U33qQe+OPfup1BE4Rt886/bNIr0TzfDh
# 5uUzoL485HjVG8wg8jBzsCIc9oTWm1wAAuEoUkv/EktA6u6wGgYGnoTm5/DbhEb7
# c9krQrbJVzTHFsCm6yG5qg73/tvK67wXy7hn6+M+T9uplIZkVckJCsDZBHFKEUta
# ZMO8eHitTEcmZQeZ1c02YKEzU7P2eyrViUA8caWr+JlZ/eObkkvdBb0LDHgGK89T
# 2L0SmlsnhoU/kb7geIBzVN+nHWcrarauTYmAJAhScFDzAf9Eri+a4OFJCOHhW9c4
# 0Z4Kip2UJ5vKo7nb4jZq42+5WGLgNng2AfrBp4l6JlOjXLvSsuuKy2MIL/4e81Yp
# 4jWb2P/ppb1tS1ksiSwvUru1KZDaQ0e8ct282b+Awdywq7RLHVg2N2Trm+GFF5op
# ov3mCNKS/6D4fOHpp9Ewjl8mUCvHouKXd4rv2E0+JuuZQGDzPGcMtghyKTVTgTTc
# MIIG9TCCBN2gAwIBAgIMeWPZY2rjO3HZBQJuMA0GCSqGSIb3DQEBCwUAMFkxCzAJ
# BgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMS8wLQYDVQQDEyZH
# bG9iYWxTaWduIEdDQyBSNDUgQ29kZVNpZ25pbmcgQ0EgMjAyMDAeFw0yMzAzMjcx
# MDIxMzRaFw0yNjAzMjMxNjE4MThaMGMxCzAJBgNVBAYTAkRLMRAwDgYDVQQHEwdL
# b2xkaW5nMRAwDgYDVQQKEwcybGlua0lUMRAwDgYDVQQDEwcybGlua0lUMR4wHAYJ
# KoZIhvcNAQkBFg9tb2tAMmxpbmtpdC5uZXQwggIiMA0GCSqGSIb3DQEBAQUAA4IC
# DwAwggIKAoICAQDMpI1rTOoWOSET3lSFQfsl/t83DCUEdoI02fNS5xlURPeGZNhi
# xQMKrhmFrdbIaEx01eY+hH9gF2AQ1ZDa7orCVSde1LDBnbFPLqcHWW5RWyzcy8Pq
# gV1QvzlFbmvTNHLm+wn1DZJ/1qJ+A+4uNUMrg13WRTiH0YWd6pwmAiQkoGC6FFwE
# usXotrT5JJNcPGlxBccm8su3kakI5B6iEuTeKh92EJM/km0pc/8o+pg+uR+f07Pp
# WcV9sS//JYCSLaXWicfrWq6a7/7U/vp/Wtdz+d2DcwljpsoXd++vuwzF8cUs09uJ
# KtdyrN8Z1DxqFlMdlD0ZyR401qAX4GO2XdzH363TtEBKAwvV+ReW6IeqGp5FUjnU
# j0RZ7NPOSiPr5G7d23RutjCHlGzbUr+5mQV/IHGL9LM5aNHsu22ziVqImRU9nwfq
# QVb8Q4aWD9P92hb3jNcH4bIWiQYccf9hgrMGGARx+wd/vI+AU/DfEtN9KuLJ8rNk
# LfbXRSB70le5SMP8qK09VjNXK/i6qO+Hkfh4vfNnW9JOvKdgRnQjmNEIYWjasbn8
# GyvoFVq0GOexiF/9XFKwbdGpDLJYttfcVZlBoSMPOWRe8HEKZYbJW1McjVIpWPnP
# d6tW7CBY2jp4476OeoPpMiiApuc7BhUC0VWl1Ei2PovDUoh/H3euHrWqbQIDAQAB
# o4IBsTCCAa0wDgYDVR0PAQH/BAQDAgeAMIGbBggrBgEFBQcBAQSBjjCBizBKBggr
# BgEFBQcwAoY+aHR0cDovL3NlY3VyZS5nbG9iYWxzaWduLmNvbS9jYWNlcnQvZ3Nn
# Y2NyNDVjb2Rlc2lnbmNhMjAyMC5jcnQwPQYIKwYBBQUHMAGGMWh0dHA6Ly9vY3Nw
# Lmdsb2JhbHNpZ24uY29tL2dzZ2NjcjQ1Y29kZXNpZ25jYTIwMjAwVgYDVR0gBE8w
# TTBBBgkrBgEEAaAyATIwNDAyBggrBgEFBQcCARYmaHR0cHM6Ly93d3cuZ2xvYmFs
# c2lnbi5jb20vcmVwb3NpdG9yeS8wCAYGZ4EMAQQBMAkGA1UdEwQCMAAwRQYDVR0f
# BD4wPDA6oDigNoY0aHR0cDovL2NybC5nbG9iYWxzaWduLmNvbS9nc2djY3I0NWNv
# ZGVzaWduY2EyMDIwLmNybDATBgNVHSUEDDAKBggrBgEFBQcDAzAfBgNVHSMEGDAW
# gBTas43AJJCja3fTDKBZ3SFnZHYLeDAdBgNVHQ4EFgQUMcaWNqucqymu1RTg02YU
# 3zypsskwDQYJKoZIhvcNAQELBQADggIBAHt/DYGUeCFfbtuuP5/44lpR2wbvOO49
# b6TenaL8TL3VEGe/NHh9yc3LxvH6PdbjtYgyGZLEooIgfnfEo+WL4fqF5X2BH34y
# EAsHCJVjXIjs1mGc5fajx14HU52iLiQOXEfOOk3qUC1TF3NWG+9mezho5XZkSMRo
# 0Ypg7Js2Pk3U7teZReCJFI9FSYa/BT2DnRFWVTlx7T5lIz6rKvTO1qQC2G3NKVGs
# HMtBTjsF6s2gpOzt7zF3o+DsnJukQRn0R9yTzgrx9nXYiHz6ti3HuJ4U7i7ILpgS
# RNrzmpVXXSH0wYxPT6TLm9eZR8qdZn1tGSb1zoIT70arnzE90oz0x7ej1fC8IUA/
# AYhkmfa6feI7OMU5xnsUjhSiyzMVhD06+RD3t5JrbKRoCgqixGb7DGM+yZVjbmhw
# cvr3UGVld9++pbsFeCB3xk/tcMXtBPdHTESPvUjSCpFbyldxVLU6GVIdzaeHAiBy
# S0NXrJVxcyCWusK41bJ1jP9zsnnaUCRERjWF5VZsXYBhY62NSOlFiCNGNYmVt7fi
# b4V6LFGoWvIv2EsWgx/uR/ypWndjmV6uBIN/UMZAhC25iZklNLFGDZ5dCUxLuoyW
# PVCTBYpM3+bN6dmbincjG0YDeRjTVfPN5niP1+SlRwSQxtXqYoDHq+3xVzFWVBqC
# NdoiM/4DqJUBMYIDCjCCAwYCAQEwaTBZMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQ
# R2xvYmFsU2lnbiBudi1zYTEvMC0GA1UEAxMmR2xvYmFsU2lnbiBHQ0MgUjQ1IENv
# ZGVTaWduaW5nIENBIDIwMjACDHlj2WNq4ztx2QUCbjAJBgUrDgMCGgUAoHgwGAYK
# KwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIB
# BDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU
# CNyZlKF0E2vkuoTccE8DQ6iGnXIwDQYJKoZIhvcNAQEBBQAEggIAcVGnkAwi/kUP
# Yr5FoXCccfOOG2m4KMfJo7qSMsKmtrqFGsoeSOBOymhCfJrj4kAVo6SVpVzkio1c
# 2YxVwDg2/N6gST7pw20KtXaHvqvWNbB0g8C/FCmsM7yryvHs5cHqUBTDxkcrDdnZ
# PkWWQ5o0LCJJSNANsijw799Fzsde5+l1XEbwBESLTBGUpyZanZd3/ROt6yKO7MWl
# kps5bmShCiBICpuPprIMuK7L0AYjMvH3SLM28qVTfKnE82hTauU58V92cZRgyLQA
# ZZT8qQZCeMkinIJFsD05htnYhpPaIHCBxg4xtAAb93LulXGHrxBMYV+llyB/dZTL
# rTn6AkZM+/g/E1+5gV47DD5v/380VlF/mtsyjwlkq8mpBRDiKp7zUoJqrkg0/Xnn
# QBocj89BWFv7G2ZOTdGiJj74EBgOuisqlMyI88Vx+U9iNFU0z4Wg1S1prGf2LZIB
# 9V/8ly5SG/pDPxyeUVODEXxJSLRk8ZFTop1HU1Ad8sr8SAwecYTsi0pbNp1ZpxaT
# h3Gj7BT4NBqphiI2XP6/cd1VlQQmOEcNDCiLsA96Rey6I4giDVnkGZ7pQIJBVLp3
# TgnJRLuDjzeYJ3EryPccKHH8cMHl3vFR2B+tWYYA1rUu/ryOs0gbILBqnYn4DSWg
# W+okTS8QZdBhh3vBu9plMzI5PtKww7k=
# SIG # End signature block

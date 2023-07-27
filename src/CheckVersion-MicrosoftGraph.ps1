Function CheckVersion-Microsoft.Graph
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
    CheckVersion-Microsoft.Graph

    # Show details of installed Microsoft.Graph including version details
    CheckVersion-Microsoft.Graph -ShowVersionDetails

    # Show details of installed Microsoft.Graph and install latest (if found)
    CheckVersion-Microsoft.Graph -InstallLatestMicrosoftGraph

    # Show details of installed Microsoft.Graph and clean-up old versions (if found)
    CheckVersion-Microsoft.Graph -CleanupOldMicrosoftGraphVersions

    # Show details, install latest (if found) and clean-up old versions (if found)
    CheckVersion-Microsoft.Graph -InstallLatestMicrosoftGraph -CleanupOldMicrosoftGraphVersions

#>
    [CmdletBinding()]
    param(
            [parameter()]
                [ValidateSet("CurrentUser","AllUsers")]
                $Scope = "AllUsers",
            [Parameter()]
                [switch]$CleanupOldMicrosoftGraphVersions = $false,
            [Parameter()]
                [switch]$InstallLatestMicrosoftGraph = $False,
            [Parameter()]
                [switch]$ShowVersionDetails = $False
         )

    #-----------------------------------------------------------------------------------------

    Write-Host "[ 1 ] Checking if Microsoft Graph is installed"
        $Installed = Get-module Microsoft.Graph -ListAvailable

        If ($Installed)
            {    
                Write-host ""
                Write-Host "      OK - Microsoft Graph was detected"
            }
        Else
            {
                Write-host ""
                Write-Host "      Microsoft.Graph was not found on this computer - installing now !"
                Install-module Microsoft.Graph -Scope $Scope -Force
            }

    #-----------------------------------------------------------------------------------------
        Write-host ""
        Write-Host "[ 2 ] Checking if you are running latest version of Microsoft.Graph from PSGallery"

        $InstalledVersions = Get-module Microsoft.Graph -ListAvailable
        $InstalledVersionsCount = ($InstalledVersions | Measure-Object).count
        $LatestVersion = $InstalledVersions | Sort-Object Version -Descending | Select-Object -First 1

        If ($ShowVersionDetails)
            {
                Write-host ""
                Write-Host "      Installed versions of Microsoft.Graph are [ $($InstalledVersionsCount) ]"
                $InstalledVersions | Select-Object Name, Version | ft
            }

        # Checking latest version in PSGallery
        $online = Find-Module -Name Microsoft.Graph -Repository PSGallery

        #compare versions
        $NewerVersionDetected = $false   # default
        if ( ([version]$online.version) -gt ([version]$LatestVersion.version) ) 
            {
                $NewerVersionDetected = $true
                Write-host ""
                Write-host "      Newer version ($($online.version)) of Microsoft.Graph was detected in PSGallery"
            }
        Else
            {
                # No new version detected ... continuing !
                Write-host ""
                Write-host "      OK - Running latest version of Microsoft.Graph"
                $NewerVersionDetected = $false
            }

    #-----------------------------------------------------------------------------------------

        Write-host ""
        Write-Host "[ 3 ] Checking if you have any older versions of Microsoft Graph installed that may conflict and should be removed"

        $VersionsCleanup = $InstalledVersions | Where-Object { [version]$_.Version -lt [version]$Online.Version }
        $VersionsCleanupCount = ($VersionsCleanup | Measure-Object).count

        Write-Host ""
        Write-Host "      You have $VersionsCleanupCount older versions of Microsoft.Graph to remove"
        Write-Host ""

        If ($ShowVersionDetails)
            {
                $VersionsCleanup | Select-Object Version
            }

    #-----------------------------------------------------------------------------------------

    # Update to latest version of Microsoft Graph
        If ( ($InstallLatestMicrosoftGraph) -and ($NewerVersionDetected -eq $true) )
            {
                Write-host ""
                Write-host "Updating to latest version $($online.version) of Microsoft.Graph from PSGallery ... Please Wait !"
                Update-module Microsoft.Graph -Force
            }
        ElseIf ( ($InstallLatestMicrosoftGraph) -and ($NewerVersionDetected -eq $false) )
            {
                Write-host ""
                Write-host "OK - Running latest version of Microsoft.Graph"
            }

    #-----------------------------------------------------------------------------------------

    # Remove old versions of Microsoft.Graph

    If ( ($CleanupOldMicrosoftGraphVersions) -and ($VersionsCleanupCount -gt 0) )
        {
            ForEach ($ModuleRemove in $VersionsCleanup)
                {
                    Write-Host "Removing old version $($ModuleRemove.Version) of Microsoft.Graph ... Please Wait !"
                    Uninstall-module Microsoft.Graph -RequiredVersion $ModuleRemove.Version -Force
                }
        }
}

# SIG # Begin signature block
# MIIRgwYJKoZIhvcNAQcCoIIRdDCCEXACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQULR7G06f18ewiaEay14BQszpU
# +E2ggg3jMIIG5jCCBM6gAwIBAgIQd70OA6G3CPhUqwZyENkERzANBgkqhkiG9w0B
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
# jTE/0lRa56As8+VPqaxM8gdbHj0wDQYJKoZIhvcNAQEBBQAEggIAqEaxqymfsrdb
# 98NEc1qBxN4ab9x1a7J4zl4gvwoWg3rvS3pxkk15OsOBsxqGjkA1bnn4FimQCyHK
# EIK8yNrc+3P0p6fha60evpV/Bv6Lk6Y7QfTsDlqiblO9/HdC1eDBNspNsUM1u/NZ
# Ztl/U2kceBY2tyufp+Z2ScBvTwdl5m7zSXMrLd69odlM83D6TXuFRF07iDxQHIp9
# GWS+4UwXLDcip1UaCYcoN5ZE0WtDuDQEbsRJ8A75cpLechLWJ10PonO/mBzvnZNU
# 58Ygp3ADEM0RVF6Zb1FNn8XtXbJ1ihwmd3hrkGpvf4FJNSlS2dahb4vnX+iNgY71
# 0gCQ4OMIASQ0pmy6SUpKvQNXBoVBw13xIQdWxT8NDvg0e0zOikHB11fy4rdQ5rZc
# io0x5JZbLwBQp8U988wxoLQ2gehqpjC9llCn8/ePLZXEOEXiFBIw1Yutrsnnyo+B
# nZ1ZyuVmRRiuDZf1Bsp6kzgmF0QGmhYVR6v9lUnqITnaQGHdqmAYXSYSQAPPuN+y
# oVBUYPAyG+nauUfGC/o2VdUo8mudtnvqF5TUeAbf9DB6nrIthUi4+fq+B0i5Z5+k
# kX2CIqQsOddHKLRO8TnY3YOotinh9rmaSwPdWcrSKJuvy7h4hvJj9fYdi1Q3gpY9
# OdIrejuiU41JF9iTZlPX1fw4hCE56Gk=
# SIG # End signature block

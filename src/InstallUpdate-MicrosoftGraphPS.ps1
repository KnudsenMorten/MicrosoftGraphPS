Function InstallUpdate-MicrosoftGraphPS {
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

    #####################################################################
    # MicrosoftGraphPS
    #####################################################################
    $Module = "MicrosoftGraphPS"

    $ModuleCheck = Get-Module -Name $Module -ListAvailable -ErrorAction SilentlyContinue
        If (!($ModuleCheck))
            {
                Write-host ""
                Write-host "Installing latest version of $($Module) from PsGallery in scope $($Scope) .... Please Wait !"

                Install-module -Name $Module -Repository PSGallery -Force -Scope $Scope
                import-module -Name $Module -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
            }
        Else
            {
                #####################################
                # Check for any available updates                    
                #####################################

                    # Current version
                    $InstalledVersions = Get-module $Module -ListAvailable

                    $LatestVersion = $InstalledVersions | Sort-Object Version -Descending | Select-Object -First 1

                    $CleanupVersions = $InstalledVersions | Where-Object { $_.Version -ne $LatestVersion.Version }

                    # Online version in PSGallery (online)
                    $Online = Find-Module -Name $Module -Repository PSGallery

                    # Compare versions
                    if ( ([version]$Online.Version) -gt ([version]$LatestVersion.Version) ) 
                        {
                            Write-host ""
                            Write-host "Newer version ($($Online.version)) of $($Module) was detected in PSGallery"
                            Write-host ""
                            Write-host "Updating to latest version $($Online.version) of $($Module) from PSGallery ... Please Wait !"
                            
                            Update-module $Module -Force
                            import-module -Name $Module -Global -force -DisableNameChecking  -WarningAction SilentlyContinue
                        }
                    Else
                        {
                            # No new version detected ... continuing !
                            Write-host ""
                            Write-host "OK - Running latest version ($($LatestVersion.version)) of $($Module)"
                        }

                #####################################
                # Clean-up older versions, if found
                #####################################

                    $InstalledVersions = Get-module $Module -ListAvailable
                    $LatestVersion = $InstalledVersions | Sort-Object Version -Descending | Select-Object -First 1
                    $CleanupVersions = $InstalledVersions | Where-Object { $_.Version -ne $LatestVersion.Version }

                    Write-host ""
                    ForEach ($ModuleRemove in $CleanupVersions)
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

            } #If (!($ModuleCheck))
}



# SIG # Begin signature block
# MIIRgwYJKoZIhvcNAQcCoIIRdDCCEXACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUe3T+KdMlHf7MrsXoC+3uOVtL
# qN2ggg3jMIIG5jCCBM6gAwIBAgIQd70OA6G3CPhUqwZyENkERzANBgkqhkiG9w0B
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
# t2ciaasO5ufdrE6jn1gIT2+11P4wDQYJKoZIhvcNAQEBBQAEggIAx5h7tG00vQXe
# hcFN8s5qzLQqV7IpUUVVTh8VYWNpI1v4BUbn9r+CA2rSbdZ6QWFgKvGV98D/qLao
# BQ8q1m5VW0FQACDkBi9mdXP4K0D6IItAr4mGpUf5c8KcON4ryjPEhZc0oKVT/K4P
# uOKyL67rz2aQnVNukdg7UvGERljNYYc0+SBHs0eMLlfv5z4TWHbmaMM91JdOr5LG
# gAcR9cu7rTYdk6maNPSiLIwPf7CLDuCR+wTGe3xqS1c9Nok/MUKjAzHoELepYNAJ
# 4fxetACf1EBSp63wFKA9McKYr40B2TG0RxWNX4vUi4QLSssCzpNx42l7jm33wHKJ
# gIKkNmgIE89syi546FKqp4gKtdP+BUyy4xHEFbalUr8SQnc9QsgvaA2h+raz5Z8p
# sFNANHg1msIC0HgcFwxhwPwwKBLjjGp8Jvmz+Hi7aPv/VD1gQjDd25si00ktT8A2
# F6spOtLciO5iK+GNeWVmwegR2RVtH0i0VYxUPrFPkuu0WP07mn5J2j38VuZcTmp+
# pZ6fNv2Fqhqe0d2WjVr5mSWgIoLqZmMwgQXXJ9zbXny85Hs+zVE9M6E3HOfsY5AX
# tuTvKq0auq/iHAiCMziUbW0o3aQ9cEJVVO/Av1EhFwIZNyjv8ziyDjIbUrLZNrcn
# Hi7B0MoJg8AUSO/e19rZD06zwZWHiJM=
# SIG # End signature block

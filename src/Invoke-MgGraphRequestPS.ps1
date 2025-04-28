Function Invoke-MgGraphRequestPS {
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
                $OutputType = "PSObject",
            [Parameter()]
                [Object]$Body,
            [Parameter()]
                [Object]$Headers,
            [Parameter()]
                [Object]$ContentType,
            [Parameter()]
                [boolean]$SkipHeaderValidation,
            [Parameter()]
                [switch]$PassThru
         )

<# TROUBLESHOOTING !!
    
    $Uri = $Uri
    $Method = "GET"
    $OutputType = "PSObject"

#>
    $Result        = @()
    $ResultsCount  = 0

    $CmdToRun_Hash = @{}

    If ($Method)
        {
            $CmdToRun_Hash += @{ Method = $Method }
        }
    If ($Uri)
        {
            $CmdToRun_Hash += @{ Uri = $Uri }
        }
    If ($Headers)
        {
            $CmdToRun_Hash += @{ Headers = $Headers }
        }
    If ($Body)
        {
            $CmdToRun_Hash += @{ Body = $Body }
        }
    If ($OutputType)
        {
            $CmdToRun_Hash += @{ OutputType = $OutputType }
        }
    If ($ContentType)
        {
            $CmdToRun_Hash += @{ ContentType = $ContentType }
        }
    If ($SkipHeaderValidation)
        {
            $CmdToRun_Hash += @{ SkipHeaderValidation = $SkipHeaderValidation }
        }
    If ($PassThru)
        {
            $CmdToRun_Hash += @{ PassThru = $PassThru }
        }

    $ResultsRaw   = Invoke-MGGraphRequest @CmdToRun_Hash

    $Result       += $ResultsRaw.value
    $ResultsCount += ($ResultsRaw.value | Measure-Object).count
    Write-host "[ $($ResultsCount) ]  Getting data from $($Uri) using MgGraph"

    if (!([string]::IsNullOrEmpty($ResultsRaw.'@odata.nextLink'))) 
        { 
            do 
                { 
                    Try
                        {
                            $Uri = $ResultsRaw.'@odata.nextLink'

                            $CmdToRun_Hash = @{}

                            If ($Method)
                                {
                                    $CmdToRun_Hash += @{ Method = $Method }
                                }
                            If ($Uri)
                                {
                                    $CmdToRun_Hash += @{ Uri = $Uri }
                                }
                            If ($Headers)
                                {
                                    $CmdToRun_Hash += @{ Headers = $Headers }
                                }
                            If ($Body)
                                {
                                    $CmdToRun_Hash += @{ Body = $Body }
                                }
                            If ($OutputType)
                                {
                                    $CmdToRun_Hash += @{ OutputType = $OutputType }
                                }
                            If ($ContentType)
                                {
                                    $CmdToRun_Hash += @{ ContentType = $ContentType }
                                }
                            If ($SkipHeaderValidation)
                                {
                                    $CmdToRun_Hash += @{ SkipHeaderValidation = $SkipHeaderValidation }
                                }
                            If ($PassThru)
                                {
                                    $CmdToRun_Hash += @{ PassThru = $PassThru }
                                }

                            $ResultsRaw   = Invoke-MGGraphRequest @CmdToRun_Hash
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



# SIG # Begin signature block
# MIIRgwYJKoZIhvcNAQcCoIIRdDCCEXACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU/JY79PGVA3xJnFQ8FtTKzYZ1
# tW+ggg3jMIIG5jCCBM6gAwIBAgIQd70OA6G3CPhUqwZyENkERzANBgkqhkiG9w0B
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
# i8Sda27fSQHgvAgtSDoFr/ZJo5cwDQYJKoZIhvcNAQEBBQAEggIAfqikAn3/Rolg
# /S0bgidiSH9OdWIkVbnvM1N1gUgyRwmNe0pvM9XpEr2E9n8FEWwC96Mdm/GeTUmC
# KDjFtOuDIV5/mBGQxPPk/US4/EcB1EqDKrJfwreL8KTwcQnanmZh4R7gzDBm7INx
# 3UrO8bhuGCL2dew7i/SuVpsaLwLlam09vrjJbBbYqW0QpFZhIy3H2ClmgWa4K/nX
# yd8slTf7wM6mCVNz1+1ey0OrFtzeFtxZxSi6eiigBwGK5sqgvK24/dEFlR3Ks5po
# 2Ger6juMA2uO/lPigdbAm8DBs0+36fhiM8nTHoGCVNNb5ni/FKiH2PJD/ztmv+sq
# /BJyV6zZjNnVZcIwBLrAelZ7fH/B1nMvQWMFMjW7FH0GEfRk68yX5fwfuiZOyoLA
# v7s65XbQLuYgWNtlf1qzkJWyRVpfODNl8UsVkyGJk+gxhoeonQO/Njepo/yK/vnQ
# hH2LEu82f+wSpQzFSPsrdyMI9LQH2wtcbSVMr0TfhDYvseIuKEbB/CE4bUx9lV/L
# 2T8FE/xbJtj8DBPjfWKsp3atYThg6Pn7oqd/4ngYp5wbhMvtcq9DH3+7TDmcFhtp
# fg4pBsENTY6ZhVpJaqQf66a2rfToyOu7d9kua/PACzksrzCCT5TRMfyjcTFCY3JL
# 6dY6paRs4dTlixtjWFyNbF5NaADs9A8=
# SIG # End signature block

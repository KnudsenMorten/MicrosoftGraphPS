Function Invoke-MgGraphRequestPS
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
    # Method #1 - REST Uri call
    $Uri        = "https://graph.microsoft.com/v1.0/devicemanagement/managedDevices"
    $Devices    = Invoke-MgGraphRequestPS -Uri $Uri -Method GET -OutputType PSObject

    # Method #2 - MgGraph cmdlet
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

    $ResultsRaw   = Invoke-MGGraphRequest -Method $Method -Uri $Uri  -OutputType $OutPutType

    $Result       += $ResultsRaw.value
    $ResultsCount += ($ResultsRaw.value | Measure-Object).count
    Write-host "[ $($ResultsCount) ]  Getting data from $($Uri) using MgGraph"

    if (!([string]::IsNullOrEmpty($ResultsRaw.'@odata.nextLink'))) 
        { 
            do 
                { 
                    Try
                        {
                            $ResultsRaw    = Invoke-MGGraphRequest -Method $Method -Uri $ResultsRaw.'@odata.nextLink'  -OutputType $OutPutType
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

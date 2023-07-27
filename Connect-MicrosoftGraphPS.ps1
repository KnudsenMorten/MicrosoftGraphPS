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
    Connect-MicrosoftGraphPS -Scopes "DeviceManagementConfiguration.ReadWrite.All"
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

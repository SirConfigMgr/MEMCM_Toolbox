<#
    .SYNOPSIS
    Get Device Variables

    .DESCRIPTION
    The script queries all device variables of one or all devices and optionally exports them to a CSV file.

    .PARAMETER All
    All devices are queried

    .PARAMETER DeviceName
    Device to be queried

    .PARAMETER SiteServer
    FQDN of the MEMCM Site Server

    .PARAMETER SiteCode
    MEMCM Site Code

    .PARAMETER ExportPath
    If set, a CSV file with the queried variables will be created. 

    .EXAMPLE
    PS> Get-MECMDeviceVariables.ps1 -All -SiteServer "mecm.domain.com" -SiteCode "AA1"

    .EXAMPLE
    PS> Get-MECMDeviceVariables.ps1 -All -ExportPath "C:\DeviceVariables.csv" -SiteServer "mecm.domain.com" -SiteCode "AA1"

    .EXAMPLE
    PS> Get-MECMDeviceVariables.ps1 -DeviceName "ABC123" -SiteServer "mecm.domain.com" -SiteCode "AA1"

    .EXAMPLE
    PS> Get-MECMDeviceVariables.ps1 -DeviceName "ABC123" -ExportPath "C:\DeviceVariables.csv" -SiteServer "mecm.domain.com" -SiteCode "AA1"
#>

param (
    [parameter(Mandatory=$true,ParameterSetName="All")][Switch]$All,
    [parameter(Mandatory=$true,ParameterSetName="One")][String]$DeviceName,
    [parameter(Mandatory=$true)][String]$SiteServer,
    [parameter(Mandatory=$true)][String]$SiteCode,
    [String]$ExportPath
    )

if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
}

if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer
}

Set-Location "$($SiteCode):\"

[System.Collections.ArrayList]$DeviceVariables = @()

If ($All) {
    $Devices = (Get-CMDevice -Fast).Name
    Foreach ($Device in $Devices) {
        $Variables = Get-CMDeviceVariable -DeviceName $Device
        Foreach ($Variable in $Variables) {
            $DV = [pscustomobject]@{'Device'=$Device;'VariableName'=$Variable.Name;'VariableValue'=$Variable.Value}
            $DeviceVariables.add($DV)
            $DV = $null
            }
        }
    }

Else {
    $Variables = Get-CMDeviceVariable -DeviceName $DeviceName
    Foreach ($Variable in $Variables) {
        $DV = [pscustomobject]@{'Device'=$DeviceName;'VariableName'=$Variable.Name;'VariableValue'=$Variable.Value}
        $DeviceVariables.add($DV)
        $DV = $null
        }
    }
    
If ($ExportPath) {$DeviceVariables | Export-Csv -Path $ExportPath -Delimiter ";"}
Else {$DeviceVariables}


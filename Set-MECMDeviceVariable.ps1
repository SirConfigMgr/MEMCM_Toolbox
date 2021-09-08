<#
    .SYNOPSIS
    Set Device Variables

    .DESCRIPTION
    The script creates device variables. Either a single variable to a device or multiple variables are created by importing a file. The import file must have the following format:
    
    ComputerName;VariableName;VariableValue;Masked

    The file does not need a header. "Masked" is either true or false (default = false).

    .PARAMETER DeviceName
    Target device name.

    .PARAMETER VariableName
    Variable Name

    .PARAMETER VariableValue
    Variable Value

    .PARAMETER ImportFile
    Full path of import file (.csv)

    .PARAMETER SiteServer
    FQDN of the MEMCM Site Server

    .PARAMETER SiteCode
    MEMCM Site Code 

    .EXAMPLE
    PS> Set-MECMDeviceVariable -DeviceName "ABC001" -VariableName "Var1" -VariableValue "Value1" -VariableMasked $false -$SiteServer "memcm.domain.com" -SiteCode "AA1"

    .EXAMPLE
    PS> Set-MECMDeviceVariable -ImportFile "C:\Import.csv" -$SiteServer "memcm.domain.com" -SiteCode "AA1"
#>

param(
    [parameter(Mandatory=$true)][String]$SiteServer,
    [parameter(Mandatory=$true)][String]$SiteCode,
    [parameter(Mandatory=$true,ParameterSetName="One")][String]$DeviceName,
    [parameter(Mandatory=$true,ParameterSetName="One")][String]$VariableName,
    [parameter(Mandatory=$true,ParameterSetName="One")][String]$VariableValue,
	[parameter(Mandatory=$false,ParameterSetName="One")][Bool]$VariableMasked = $false,
    [parameter(Mandatory=$true,ParameterSetName="Import")][String]$ImportFile
	)


Function Set-Variable {

param (
    [String]$DN,
    [String]$VN,
    [String]$VV,
    [Bool]$VM = $false
    )

    $Device = Get-WmiObject -Computername $Siteserver -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_R_System WHERE Name = '$($DN)'"
    $MachineSettings = Get-WmiObject -Computername $Siteserver -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_MachineSettings WHERE ResourceID = '$($Device.ResourceID)'"
    $SMSMachineSettings = [WmiClass]"\\$($Siteserver)\ROOT\SMS\site_$($SiteCode):SMS_MachineSettings"

    If ($MachineSettings -ne $null){
        $MachineSettings.Get()
        If ($MachineSettings.MachineVariables.length -ne 0){$VariableIndex = $MachineSettings.MachineVariables.length}
        Else {$VariableIndex = 0}

        $MachineSettings.MachineVariables = $MachineSettings.MachineVariables += [WmiClass]"\\$($Siteserver)\ROOT\SMS\site_$($SiteCode):SMS_MachineVariable"
        $MachineVariables = $MachineSettings.MachineVariables
        $MachineVariables[$VariableIndex].Name=$VN
        $MachineVariables[$VariableIndex].Value=$VV
        $MachineVariables[$VariableIndex].Ismasked = $VM

        $MachineSettings.MachineVariables = $MachineVariables
        $MachineSettings.put()
        }
    Else {
        $MachineSettings = $SMSMachineSettings.CreateInstance()
        $MachineSettings.psbase.properties["ResourceID"].value = $($Device.ResourceID)
        $MachineSettings.psbase.properties["SourceSite"].value = $($SiteCode)
        $MachineSettings.psbase.properties["LocaleID"].value = 1031
        $MachineSettings.MachineVariables = $MachineSettings.MachineVariables + [WmiClass]"\\$($Siteserver)\ROOT\SMS\site_$($SiteCode):SMS_MachineVariable"
        $MachineVariables = $MachineSettings.MachineVariables
        $MachineVariables[0].Name=$VN
        $MachineVariables[0].Value=$VV
        $MachineVariables[0].Ismasked = $VM

        $MachineSettings.MachineVariables = $MachineVariables
        $MachineSettings.put()
        }
    }

If ($Device) {Set-Variable -DN $DeviceName -VN $VariableName -VV $VariableValue -VM $VariableMasked}

If ($ImportFile) {
    $Content = Import-Csv -Path $ImportFile -Delimiter ";" -Header Device,VariableName,VariableValue,VariableMasked
    Foreach ($Item in $Content) {
        [boolean]$Masked = [System.Convert]::ToBoolean($Item.VariableMasked)
        Set-Variable -DN $Item.Device -VN $Item.VariableName -VV $Item.VariableValue -VM $Masked
        }
    }

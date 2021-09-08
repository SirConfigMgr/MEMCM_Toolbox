<#
    .SYNOPSIS
    Import Computer, add to collection, set variables

    .DESCRIPTION
    The script imports computers into MECM, adds them to a collection and sets variables. Three CSV files are read for the import, respectively for importing the computers (import-computer.csv), adding them to a collection (add-collection.csv) and setting the variables (set-variable.csv). 

    .PARAMETER ImportFolder
    Folder with the three required files. (Full path)

    .PARAMETER SiteServer
    FQDN of the MEMCM Site Server

    .PARAMETER SiteCode
    MEMCM Site Code 

    .EXAMPLE
    PS> Import-MECMComputerFromFolder -ImportFolder "C:\Folder" -$SiteServer "memcm.domain.com" -SiteCode "AA1"

#>

param (
    [parameter(Mandatory=$true)][String]$SiteServer,
    [parameter(Mandatory=$true)][String]$SiteCode,
    [parameter(Mandatory=$true)][String]$ImportFolder
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
        $Error.clear()
        $MachineSettings.put() | Out-Null
        If ($Error) {Write-Host -ForegroundColor White -BackgroundColor DarkRed "Failed To Set Variable $VN Value $VV On $DN"}
        Else {Write-Host -ForegroundColor White -BackgroundColor DarkGreen "Set Variable $VN Value $VV On $DN"}
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
        $Error.clear()
        $MachineSettings.put() | Out-Null
        If ($Error) {Write-Host -ForegroundColor White -BackgroundColor DarkRed "Failed To Set Variable $VN Value $VV On $DN"}
        Else {Write-Host -ForegroundColor White -BackgroundColor DarkGreen "Set Variable $VN Value $VV On $DN"}
        }
    }

Function New-Device {

param (
    [String]$DN,
    [String]$MAC
    )

$OldDevice = Get-WmiObject -Computername $Siteserver -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_R_System WHERE Name = '$($DN)'"
If ($OldDevice) {
    $Error.clear()
    $OldDevice.psbase.delete() | Out-Null
    If ($Error) {Write-Host -ForegroundColor White -BackgroundColor DarkRed "Failed To Delete Old $OldDevice"}
    Else {Write-Host -ForegroundColor White -BackgroundColor DarkGreen "Deleted Old $OldDevice"}
    }
$Error.clear()
Invoke-WmiMethod -Namespace root/SMS/site_$($SiteCode) -Class SMS_Site -Name ImportMachineEntry -ArgumentList @($null, $null, $null, $null, $null, $null, $MAC, $null, $DN, $True, $null, $null) -ComputerName $SiteServer | Out-Null
If ($Error) {Write-Host -ForegroundColor White -BackgroundColor DarkRed "Failed To Create $DN"}
Else {Write-Host -ForegroundColor White -BackgroundColor DarkGreen "Device $DN Created"}
    }

Function Add-CollectionMembership {

param (
    [parameter(Mandatory=$true)][String]$DN,
    [parameter(Mandatory=$true)][String]$CN
    )

$Device = Get-WmiObject -Computername $Siteserver -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_R_System WHERE Name = '$($DN)'"
$Collection = Get-WmiObject -Computername $Siteserver -Namespace "root\sms\site_$SiteCode" -Query "SELECT * FROM SMS_Collection WHERE Name = '$($CN)'"

$MembershipRule = ([WMIClass]”\\$SiteServer\root\sms\site_$($SiteCode):SMS_CollectionRuleDirect”).CreateInstance()
$MembershipRule.ResourceClassName = "SMS_R_System"
$MembershipRule.RuleName = $Device.Name
$MembershipRule.ResourceID = $Device.ResourceID

$Error.clear()
$Collection.AddMembershipRule($MembershipRule) | Out-Null
$Collection.Put() | Out-Null
If ($Error) {Write-Host -ForegroundColor White -BackgroundColor DarkRed "Failed To Add $DN To $CN"}
Else {Write-Host -ForegroundColor White -BackgroundColor DarkGreen "$DN Added To $CN"}
$Collection.RequestRefresh() | Out-Null

    }

$Content_ImportComputerCSV = Import-Csv -Path "$ImportFolder\import_computer.csv" -Delimiter ";" -Header Device,MAC
$Content_AddCollectionCSV = Import-Csv -Path "$ImportFolder\add_collection.csv" -Delimiter ";" -Header Device,CollectionName
$Content_SetVariableCSV = Import-Csv -Path "$ImportFolder\set_variable.csv" -Delimiter ";" -Header Device,VariableName,VariableValue,VariableMasked

If ($Content_ImportComputerCSV) {
    Foreach ($IC in $Content_ImportComputerCSV) {
        New-Device -DN $IC.Device -MAC $IC.MAC
        }
    }

If ($Content_AddCollectionCSV) {
    Foreach ($AC in $Content_AddCollectionCSV) {
        Add-CollectionMembership -DN $AC.Device -CN $AC.CollectionName 
        }
    }

If ($Content_SetVariableCSV) {
    Foreach ($SV in $Content_SetVariableCSV) {
        [boolean]$Masked = [System.Convert]::ToBoolean($SV.VariableMasked)
        Set-Variable -DN $SV.Device -VN $SV.VariableName -VV $SV.VariableValue -VM $Masked
        }
    }

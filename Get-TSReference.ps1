
<#
    .SYNOPSIS
    Get task sequence references for an object

    .NOTES
    Version         1.0
    Author          Rene Hartmann
    Creation Date   02.01.2024

    .DESCRIPTION
    The script queries task sequence references for an object and list relevant information

    .PARAMETER Name
    Object name to query

    .PARAMETER ID
    Object ID to query

    .PARAMETER SiteServer
    FQDN of the MEMCM Site Server

    .PARAMETER SiteCode
    MEMCM Site Code

    .PARAMETER ExportPath
    If set, a CSV file with the queried variables will be created. 

    .EXAMPLE
    PS> Get-TSReference.ps1 -ID "AA100351" -SiteServer "mecm.domain.com" -SiteCode "AA1" -ExportPath "C:\References.csv"

    .EXAMPLE
    PS> Get-TSReference.ps1 -Name "Firefox" -SiteServer "mecm.domain.com" -SiteCode "AA1" -ExportPath "C:\References.csv"

    .EXAMPLE
    PS> Get-TSReference.ps1 -ID "AA100351" -SiteServer "mecm.domain.com" -SiteCode "AA1"

    .EXAMPLE
    PS> Get-TSReference.ps1 -Name "Firefox" -SiteServer "mecm.domain.com" -SiteCode "AA1"
#>

param (
    [parameter(Mandatory=$true,ParameterSetName="Name")]
    [String]$Name,
    [parameter(Mandatory=$true,ParameterSetName="ID")]
    [String]$ID,
    [parameter(Mandatory=$true)]
    [String]$SiteServer,
    [parameter(Mandatory=$true)]
    [String]$SiteCode,
    [String]$ExportPath
    )

If ($ID) {
    $References = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -query "SELECT * FROM SMS_TasksequencePackagereference WHERE RefpackageID = '$($ID)'" 
    }

If ($Name) {
    $References = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -query "SELECT * FROM SMS_TasksequencePackagereference WHERE ObjectName LIKE '%$($Name)%'" 
    }

[System.Collections.ArrayList]$ReferenceList = @()
Foreach ($Reference in $References) {
    If ($Reference.ObjectType -eq 0) {$ObjectType = "Package"}
    Elseif ($Reference.ObjectType -eq 3) {$ObjectType = "Driver"}
    Elseif ($Reference.ObjectType -eq 5) {$ObjectType = "Softwareupdate"}
    Elseif ($Reference.ObjectType -eq 257) {$ObjectType = "Image"}
    Elseif ($Reference.ObjectType -eq 258) {$ObjectType = "Boot Image"}
    Elseif ($Reference.ObjectType -eq 259) {$ObjectType = "OS Install Image"}
    Elseif ($Reference.ObjectType -eq 512) {$ObjectType = "Application"}
    $TSInfo = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -query "SELECT * FROM SMS_TasksequencePackage WHERE PackageID = '$($Reference.packageid)'"
    $NewReferenceObject = [pscustomobject]@{'ObjectName'=$Reference.ObjectName;'ObjectType'=$ObjectType;'TSName'=$TSInfo.Name;'TSPath'=$TSInfo.ObjectPath;'TSID'=$TSInfo.PackageID}
    $ReferenceList.add($NewReferenceObject) | out-null
    $NewReferenceObject = $null
   
}

If ($ExportPath) {$ReferenceList | Export-Csv -Path $ExportPath -Delimiter ";" -NoTypeInformation}
Else {$ReferenceList | Format-Table}
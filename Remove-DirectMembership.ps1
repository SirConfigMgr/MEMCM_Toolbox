<#
.SYNOPSIS
    This script removes the direct memberships of a specified collection (User or Device).
.PARAMETER SiteServer
    FQDN of the MEMCM Site Server
.PARAMETER SiteCode
    MEMCM Site Code
.PARAMETER CollectionID
    The ID of the collection from which the direct memberships are to be removed.
.PARAMETER CollectionName
    The Name of the collection from which the direct memberships are to be removed.
.PARAMETER All
    Remove all direct memberships rules.
.Parameter ResourceName
    Computer name or username of the member to be removed.
.Parameter Log
    Specify log path (optional).
.Parameter Silent
    Suppresses messages in the console
.EXAMPLE
    PS> Remove-DirectMembership.ps1 -SiteServer "memcm.domain.com" -SiteCode "AA1" -All -CollectionName "Device Collection X" -Log "C:\remove_membership.log"
.EXAMPLE
    PS> Remove-DirectMembership.ps1 -SiteServer "memcm.domain.com" -SiteCode "AA1" -ResourceName "Device1" -CollectionName "Device Collection X" -Log "C:\remove_membership.log"
.EXAMPLE
    PS> Remove-DirectMembership.ps1 -SiteServer "memcm.domain.com" -SiteCode "AA1" -ResourceName "Device1" -CollectionID "167223456"
.EXAMPLE
    PS> Remove-DirectMembership.ps1 -SiteServer "memcm.domain.com" -SiteCode "AA1" -ResourceName "Device1" -CollectionID "167223456" -Silent
.LINK
    https://sirconfigmgr.de/
    https://github.com/SirConfigMgr/MEMCM_Toolbox
#>

param (
    [parameter(Mandatory=$true,ParameterSetName="AllCollectionName")]
    [parameter(Mandatory=$true,ParameterSetName="AllCollectionID")]
    [Switch]$All,
    [parameter(Mandatory=$true,ParameterSetName="ResourceNameCollectionName")]
    [parameter(Mandatory=$true,ParameterSetName="ResourceNameCollectionID")]
    [String]$ResourceName,
    [parameter(Mandatory=$true,ParameterSetName="AllCollectionID")]
    [parameter(Mandatory=$true,ParameterSetName="ResourceNameCollectionID")]
    [String]$CollectionID,
    [parameter(Mandatory=$true,ParameterSetName="AllCollectionName")]
    [parameter(Mandatory=$true,ParameterSetName="ResourceNameCollectionName")]
    [String]$CollectionName,
    [parameter(Mandatory=$true)]
    [String]$SiteServer,
    [parameter(Mandatory=$true)]
    [String]$SiteCode,
    [String]$Log,
    [Switch]$Silent
    )

if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
}

if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer
}

Set-Location "$($SiteCode):\"

If ($Log) { "Start " + (Get-Date -Format "dd.MM.yyyy - hh:mm")  | Out-File -FilePath $Log -Append}

If ($ResourceName) {  
    If ($CollectionID) {$Collection = Get-CMCollection -Id $CollectionID}
    If ($CollectionName) {$Collection = Get-CMCollection -Name $CollectionName}

    $MembershipRule = Get-CMCollectionDirectMembershipRule -CollectionID $Collection.CollectionID -ResourceName $ResourceName
    If ($MembershipRule) {
        Try {
            Remove-CMCollectionDirectMembershipRule -CollectionId $Collection.CollectionID -ResourceName $ResourceName -Force
            If (!($Silent -eq $True)) {Write-Host "Removed $ResourceName From $($Collection.CollectionID) - $($Collection.Name)" -ForegroundColor Green}
            If ($Log) {"Removed $ResourceName From $($Collection.CollectionID) - $($Collection.Name)" | Out-File -FilePath $Log -Append}
            }
        Catch {
            If (!($Silent -eq $True)) {Write-Host "Failed To Removed $ResourceName From $($Collection.CollectionID) - $($Collection.Name)" -ForegroundColor Red}
            If ($Log) {"Failed To Removed $ResourceName From $($Collection.CollectionID) - $($Collection.Name)" | Out-File -FilePath $Log -Append}
            }
        }
    Else {
        If (!($Silent -eq $True)) {Write-Host "$ResourceName Is Not A Member Of $($Collection.CollectionID) - $($Collection.Name)" -ForegroundColor Red}
        If ($Log) {"$ResourceName Is Not A Member Of $($Collection.CollectionID) - $($Collection.Name)" | Out-File -FilePath $Log -Append}
        }
    }

Else {

    If ($CollectionID) {$Collection = Get-CMCollection -Id $CollectionID}
    If ($CollectionName) {$Collection = Get-CMCollection -Name $CollectionName}
    Try {
        $MemberArray = Get-CMCollectionDirectMembershipRule -CollectionId $CollectionID
        Foreach ($Member in $MemberArray) {
            Remove-CMCollectionDirectMembershipRule -CollectionId $CollectionID -ResourceId $Member.ResourceID -Force
            If (!($Silent -eq $True)) {Write-Host "Removed $($Member.RuleName) From $($Collection.CollectionID) - $($Collection.Name)" -ForegroundColor Green}
            If ($Log) {"Removed $($Member.RuleName) From $($Collection.CollectionID) - $($Collection.Name)" | Out-File -FilePath $Log -Append}
            }
        }
    Catch {
        If (!($Silent -eq $True)) {Write-Host "Failed To Removed $($Member.RuleName) From $($Collection.CollectionID) - $($Collection.Name)" -ForegroundColor Red}
        If ($Log) {"Failed To Removed $($Member.RuleName) From $($Collection.CollectionID) - $($Collection.Name)" | Out-File -FilePath $Log -Append}}
    }
    
If ($Log) { "End " + (Get-Date -Format "dd.MM.yyyy - hh:mm")  | Out-File -FilePath $Log -Append}

<#
    .SYNOPSIS
    Matches the membership rules of two computers.
    

    .DESCRIPTION
    The script queries all direct membership rules of a computer and adds the new computer to the same collections.

    .PARAMETER NewDevice
    Hostname of the new device

    .PARAMETER OldDevice
    Hostname of the old device

    .PARAMETER SiteServer
    FQDN of the MEMCM Site Server

    .PARAMETER SiteCode
    MEMCM Site Code

    .EXAMPLE
    PS> Equalize-MECMDirectCollectionMembershipRules.ps1 -NewDevice "ABC" -OldDevice "DEF" -$SiteServer "memcm.domain.com" -SiteCode "AA1"
#>

param (
    [Parameter(Mandatory=$true)][String]$NewDevice,
    [Parameter(Mandatory=$true)][String]$OldDevice,
    [Parameter(Mandatory=$true)][String]$SiteServer,
    [Parameter(Mandatory=$true)][String]$SiteCode
    )
    $NewDeviceResourceID = (Get-CMDevice -Name $NewDevice).ResourceID
    $Collections = (Get-WmiObject -ComputerName $SiteServer  -Namespace root/SMS/site_$SiteCode -Query "SELECT SMS_Collection.* FROM SMS_FullCollectionMembership, SMS_Collection where name = '$OldDevice' and SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID").Name
    Foreach ($Collection in $Collections) {
        $DirectCollectionMembership = Get-CMDeviceCollectionDirectMembershipRule -ResourceName $Device -CollectionName $Collection
        If ($DirectCollectionMembership) {Add-CMDeviceCollectionDirectMembershipRule -CollectionName $Collection -ResourceId $NewDeviceResourceID} 
        }



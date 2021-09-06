<#
    .SYNOPSIS
    Queries all packages and checks distribution to DP Groups.
    

    .DESCRIPTION
    The script checks all packages (normal packages, driver packages, OS images, OS installers, boot images, applications) 
    if they have been distributed to a DP group and exports the results to a CSV file.

    .PARAMETER Export
    Full path to Export-File (.csv) 

    .PARAMETER SiteServer
    FQDN of the MEMCM Site Server

    .PARAMETER SiteCode
    MEMCM Site Code
    
    .EXAMPLE
    PS> Get-MECMPackagesOnDPGroups -$SiteServer "memcm.domain.com" -SiteCode "AA1" -Export "C:\Packages.csv" 
#>


param (
    [Parameter(Mandatory=$true)][String]$SiteServer,
    [Parameter(Mandatory=$true)][String]$SiteCode,
    [Parameter(Mandatory=$true)][String]$Export
    )

    [System.Collections.ArrayList]$PackageStatus = @()
    $Packages = Get-CMPackage -Fast
    $Drivers = Get-CMDriverPackage -Fast
    $OSImages = Get-CMOperatingSystemImage
    $OSInstallers = Get-CMOperatingSystemInstaller
    $Bootimages = Get-CMBootImage
    $DPGroupContent = Get-WmiObject -Class SMS_DPGroupContentInfo -ComputerName "$SiteServer" -Namespace "ROOT\SMS\site_$SiteCode"
    $DPGroups = Get-CMDistributionPointGroup
    $Applications = Get-CMApplication

    Foreach ($Package in $Packages) {
        $DPGroup = $DPGroupContent -match $Package.PackageID
        If ($DPGroup) {
            $DPGroupName = ($DPGroups | Where-Object -FilterScript {$_.GroupID -eq $DPGroup.GroupID}).Name
            $Matches
            $PS = [pscustomobject]@{'Package'=$Package.Name;'PackageID'=$Package.PackageID;'PackageType'="Package";'Status'="Distributed";'DPGroup'="$DPGroupName"}
            $PackageStatus.add($PS)
            $PS = $null
            }
        Else {
            $PS = [pscustomobject]@{'Package'=$Package.Name;'PackageID'=$Package.PackageID;'PackageType'="Package";'Status'="NOT Distributed";'DPGroup'=""}
            $PackageStatus.add($PS)
            $PS = $null
            }
        }

    Foreach ($Driver in $Drivers) {
        $DPGroup = $DPGroupContent -match $Driver.PackageID
        If ($DPGroup) {
            $DPGroupName = ($DPGroups | Where-Object -FilterScript {$_.GroupID -eq $DPGroup.GroupID}).Name
            $PS = [pscustomobject]@{'Package'=$Driver.Name;'PackageID'=$Driver.PackageID;'PackageType'="DriverPackage";'Status'="Distributed";'DPGroup'="$DPGroupName"}
            $PackageStatus.add($PS)
            $PS = $null
            }
        Else {
            $PS = [pscustomobject]@{'Package'=$Driver.Name;'PackageID'=$Driver.PackageID;'PackageType'="DriverPackage";'Status'="NOT Distributed";'DPGroup'=""}
            $PackageStatus.add($PS)
            $PS = $null
            }
        }

    Foreach ($OSImage in $OSImages) {
        $DPGroup = $DPGroupContent -match $OSImage.PackageID
        If ($DPGroup) {
            $DPGroupName = ($DPGroups | Where-Object -FilterScript {$_.GroupID -eq $DPGroup.GroupID}).Name
            $PS = [pscustomobject]@{'Package'=$OSImage.Name;'PackageID'=$OSImage.PackageID;'PackageType'="OSImage";'Status'="Distributed";'DPGroup'="$DPGroupName"}
            $PackageStatus.add($PS)
            $PS = $null
            }
        Else {
            $PS = [pscustomobject]@{'Package'=$OSImage.Name;'PackageID'=$OSImage.PackageID;'PackageType'="OSImage";'Status'="NOT Distributed";'DPGroup'=""}
            $PackageStatus.add($PS)
            $PS = $null
            }
        }

    Foreach ($OSInstaller in $OSInstallers) {
        $DPGroup = $DPGroupContent -match $OSInstaller.PackageID
        If ($DPGroup) {
            $DPGroupName = ($DPGroups | Where-Object -FilterScript {$_.GroupID -eq $DPGroup.GroupID}).Name
            $PS = [pscustomobject]@{'Package'=$OSInstaller.Name;'PackageID'=$OSInstaller.PackageID;'PackageType'="OSInstaller";'Status'="Distributed";'DPGroup'="$DPGroupName"}
            $PackageStatus.add($PS)
            $PS = $null
            }
        Else {
            $PS = [pscustomobject]@{'Package'=$OSInstaller.Name;'PackageID'=$OSInstaller.PackageID;'PackageType'="OSInstaller";'Status'="NOT Distributed";'DPGroup'=""}
            $PackageStatus.add($PS)
            $PS = $null
            }
        }

    Foreach ($Bootimage in $Bootimages) {
        $DPGroup = $DPGroupContent -match $Bootimage.PackageID
        If ($DPGroup) {
            $DPGroupName = ($DPGroups | Where-Object -FilterScript {$_.GroupID -eq $DPGroup.GroupID}).Name
            $PS = [pscustomobject]@{'Package'=$Bootimage.Name;'PackageID'=$Bootimage.PackageID;'PackageType'="Bootimage";'Status'="Distributed";'DPGroup'="$DPGroupName"}
            $PackageStatus.add($PS)
            $PS = $null
            }
        Else {
            $PS = [pscustomobject]@{'Package'=$Bootimage.Name;'PackageID'=$Bootimage.PackageID;'PackageType'="Bootimage";'Status'="NOT Distributed";'DPGroup'=""}
            $PackageStatus.add($PS)
            $PS = $null
            }
        }

    Foreach ($Application in $Applications) {
        $DPGroup = $DPGroupContent -match $Application.PackageID
        If ($DPGroup) {
            $DPGroupName = ($DPGroups | Where-Object -FilterScript {$_.GroupID -eq $DPGroup.GroupID}).Name
            $PS = [pscustomobject]@{'Package'=$Application.LocalizedDisplayName;'PackageID'=$Application.PackageID;'PackageType'="Application";'Status'="Distributed";'DPGroup'="$DPGroupName"}
            $PackageStatus.add($PS)
            $PS = $null
            }
        Else {
            $PS = [pscustomobject]@{'Package'=$Application.LocalizedDisplayName;'PackageID'=$Application.PackageID;'PackageType'="Application";'Status'="NOT Distributed";'DPGroup'=""}
            $PackageStatus.add($PS)
            $PS = $null
            }
        }
    
$PackageStatus | Export-Csv -Path $Export -Delimiter ";"



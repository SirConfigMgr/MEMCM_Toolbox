<#	
    .SYNOPSIS
    Install user and machine targeted software dynamic in OSD.

    .NOTES
    Created on:    2022.09.05
    Last Updated:  2023.12.14
    Version:       1.2
    Author:        Rene Hartmann
    Filename:      Invoke-DynamicAppInstall.ps1

    .DESCRIPTION
    The script is intended for execution within an MCM task sequence and creates a list of applications that are assigned to the primary users of the device and the device itself in order to install them dynamically in the "Install application" or "Install package" step.

    .PARAMETER SiteServer
    FQDN of the MEMCM Site Server

    .PARAMETER SiteCode
    MEMCM Site Code

    .PARAMETER LogPath
    Path where the log file is saved

    .PARAMETER ResourceName
    Name of the computer that is currently being installed

    .PARAMETER UserApps
    Query and install user targeted apps (True/False)

    .PARAMETER MachineApps
    Query and install machine targeted apps (True/False)

    .CHANGELOG
    v1.0
    Initial creation

    v1.1
    Added Logging

    v1.2
    Added Application Groups

    .LINK
    https://sirconfigmgr.de/osd-install-dynamic-software/
    https://github.com/SirConfigMgr/MEMCM_Toolbox

#>

param (
    $SiteCode,
    $SiteServer,
    $LogPath,
    $ResourceName,
    $UserApps,
    $MachineApps
    )

Function Get-UserApps {
    $TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $UserLogName = "SMSTS_DynamicAppInstall4User.log"
    $Count = 1
    $FoundApps = @()
    Get-Date | Out-File "$LogPath\$UserLogName"
    "SiteCode: " + $SiteCode | Out-File "$LogPath\$UserLogName" -Append
    "SiteServer: " + $SiteServer | Out-File "$LogPath\$UserLogName" -Append
    "Resource: " + $ResourceName | Out-File "$LogPath\$UserLogName" -Append
    $PrimaryUsers = (Get-WmiObject -ComputerName $SiteServer -Class SMS_UserMachineRelationship -Namespace root\SMS\Site_$SiteCode -Filter “ResourceName='$ResourceName' and IsActive='1' and Types='1'”).UniqueUserName.replace(“\”,”\\”)
    "Primary Users: " + $PrimaryUsers | Out-File "$LogPath\$UserLogName" -Append
    If ($PrimaryUsers -ne $null) {        
        Foreach ($PrimaryUser in $PrimaryUsers){
            Write-Host $PrimaryUser
            "Check Primary User: " + $PrimaryUser | Out-File "$LogPath\$UserLogName" -Append
            $User = Get-WmiObject -ComputerName $SiteServer -Class SMS_R_User -Namespace root\SMS\Site_$SiteCode -Filter "UniqueUserName='$PrimaryUser'"
            $Collections = (Get-WmiObject -ComputerName $SiteServer -Namespace root\SMS\Site_$SiteCode -Class SMS_FullCollectionMembership -filter "ResourceID='$($User.ResourceId)'").collectionID
            "Collection Memberships: " + $Collections | Out-File "$LogPath\$UserLogName" -Append
            Foreach ($Collection in $Collections) {    
                $ApplicationNames = (Get-WmiObject -ComputerName $SiteServer -Class SMS_ApplicationAssignment -Namespace root/SMS/site_$SiteCode -Filter "TargetCollectionID='$Collection' and OfferTypeID='0'").ApplicationName
                $ApplicationGroupNames = (Get-WmiObject -ComputerName $SiteServer -Class SMS_ApplicationGroupAssignment -Namespace root/SMS/site_$SiteCode -Filter "TargetCollectionID='$Collection' and OfferTypeID='0'").ApplicationName
                If ($ApplicationNames) {
                    Write-Host "Found Applications:"
                    Foreach ($ApplicationName in $ApplicationNames) {
                        $FoundApps += $ApplicationName
                        "Application: " + $ApplicationName | Out-File "$LogPath\$UserLogName" -Append
                        $Id = “{0:D2}” -f $Count
                        $AppId = “APPId$Id”
                        $TSEnv.Value($AppId) = $ApplicationName
                        Write-Host "$AppId $ApplicationName"
                        $Count = $Count + 1
                        }
                    }
                If ($ApplicationGroupNames) {
                    Write-Host "Found Application Groups:"
                    Foreach ($ApplicationGroupName in $ApplicationGroupNames) {
                        Write-Host "$ApplicationGroupName"
                        "ApplicationGroup: " + $ApplicationGroupName | Out-File "$LogPath\$UserLogName" -Append
                        $AG = Get-WmiObject -ComputerName $SiteServer -Class SMS_ApplicationGroup -Namespace root/SMS/site_$SiteCode -Filter "LocalizedDisplayName='$ApplicationGroupName'"
                        If ($AG.ModelName.Count -gt 1) {$AG_Items = Get-WmiObject -ComputerName $SiteServer -Class SMS_ApplicationGroupItem -Namespace root/SMS/site_$SiteCode -Filter "ModelName='$($AG.ModelName[0])'"}
                        Else {$AG_Items = Get-WmiObject -ComputerName $SiteServer -Class SMS_ApplicationGroupItem -Namespace root/SMS/site_$SiteCode -Filter "ModelName='$($AG.ModelName)'"} 
                        Foreach ($AG_Item in $AG_Items) {
                            $AG_Item_APP = (Get-WmiObject -ComputerName $SiteServer -Class SMS_Application -Namespace root/SMS/site_$SiteCode -Filter "CI_UniqueID='$($AG_Item.Item_CIUniqueID)'").LocalizedDisplayName
                            $FoundApps += $AG_Item_APP
                            "Application: " + $AG_Item_APP | Out-File "$LogPath\$UserLogName" -Append
                            $Id = “{0:D2}” -f $Count
                            $AppId = “APPId$Id”
                            $TSEnv.Value($AppId) = $AG_Item_APP
                            Write-Host "$AppId $AG_Item_APP"
                            $Count = $Count + 1
                            }
                        }
                    }
                }
            If ($FoundApps.Count -eq 0) {
                $TSEnv.Value("SkipUserApplications") = "True"
                "No Apps Found - Skip App Install" | Out-File "$LogPath\$UserLogName" -Append
                Write-Host "No Apps Found - Skip App Install"
                Break
                }
            }
        }
    
    Else {
        $TSEnv.Value("SkipUserApplications") = "True"
        "No Primary User - Skip App Install" | Out-File "$LogPath\$UserLogName" -Append
        Write-Host "No Primary User - Skip applications"
        Break
        }
    }

Function Get-MachineApps {
    $TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $MachineLogName = "SMSTS_DynamicAppInstall4Machine.log"
    $A_Count = 1
    $P_Count = 1
    $FoundMachineApps = @()
    $FoundMachinePkgs = @()
    Get-Date | Out-File "$LogPath\$MachineLogName"
    "SiteCode: " + $SiteCode | Out-File "$LogPath\$MachineLogName" -Append
    "SiteServer: " + $SiteServer | Out-File "$LogPath\$MachineLogName" -Append
    "Resource: " + $ResourceName | Out-File "$LogPath\$MachineLogName" -Append
    $Computer = (Get-WmiObject -ComputerName $SiteServer -Namespace root\SMS\Site_$SiteCode -Class SMS_R_System -filter "Name='$ResourceName'")
    $Collections = (Get-WmiObject -ComputerName $SiteServer -Namespace root\SMS\Site_$SiteCode -Class SMS_FullCollectionMembership -filter "ResourceID='$($Computer.ResourceId)'").collectionID
    "Collection Memberships: " + $Collections | Out-File "$LogPath\$MachineLogName" -Append
    Foreach ($Collection in $Collections) {    
        $ApplicationNames = (Get-WmiObject -ComputerName $SiteServer -Class SMS_ApplicationAssignment -Namespace root/SMS/site_$SiteCode -Filter "TargetCollectionID='$Collection' and OfferTypeID='0'").ApplicationName
        $PackageDeploymentIDs = (Get-WmiObject -ComputerName $SiteServer -Class SMS_DeploymentInfo -Namespace root/SMS/site_$SiteCode -Filter "CollectionID='$Collection' and TargetSecurityTypeID='2' and (CollectionName like 'install%' or CollectionName like 'maintenance%')").DeploymentID
        $ApplicationGroupNames = (Get-WmiObject -ComputerName $SiteServer -Class SMS_ApplicationGroupAssignment -Namespace root/SMS/site_$SiteCode -Filter "TargetCollectionID='$Collection' and OfferTypeID='0'").ApplicationName
        If ($ApplicationNames) {
            Write-Host "Found Application:"
            Foreach ($ApplicationName in $ApplicationNames) {
                $FoundMachineApps += $ApplicationName
                "Application: " + $ApplicationName | Out-File "$LogPath\$MachineLogName" -Append
                $Id = “{0:D2}” -f $A_Count
                $M_AppId = “M_AppId$Id”
                $TSEnv.Value($M_AppId) = $ApplicationName
                Write-Host "$M_AppId $ApplicationName"
                $A_Count = $A_Count + 1
                }
            }
        If ($ApplicationGroupNames) {
            Write-Host "Found Application Groups:"
            Foreach ($ApplicationGroupName in $ApplicationGroupNames) {
                Write-Host "$ApplicationGroupName"
                "ApplicationGroup: " + $ApplicationGroupName | Out-File "$LogPath\$MachineLogName" -Append
                $AG = Get-WmiObject -ComputerName $SiteServer -Class SMS_ApplicationGroup -Namespace root/SMS/site_$SiteCode -Filter "LocalizedDisplayName='$ApplicationGroupName'"
                If ($AG.ModelName.Count -gt 1) {$AG_Items = Get-WmiObject -ComputerName $SiteServer -Class SMS_ApplicationGroupItem -Namespace root/SMS/site_$SiteCode -Filter "ModelName='$($AG.ModelName[0])'"}
                Else {$AG_Items = Get-WmiObject -ComputerName $SiteServer -Class SMS_ApplicationGroupItem -Namespace root/SMS/site_$SiteCode -Filter "ModelName='$($AG.ModelName)'"} 
                Foreach ($AG_Item in $AG_Items) {
                    $AG_Item_APP = (Get-WmiObject -ComputerName $SiteServer -Class SMS_Application -Namespace root/SMS/site_$SiteCode -Filter "CI_UniqueID='$($AG_Item.Item_CIUniqueID)'").LocalizedDisplayName
                    $FoundMachineApps += $AG_Item_APP
                    "Application: " + $AG_Item_APP | Out-File "$LogPath\$MachineLogName" -Append
                    $Id = “{0:D2}” -f $A_Count
                    $M_AppId = “APPId$Id”
                    $TSEnv.Value($M_AppId) = $AG_Item_APP
                    Write-Host "$M_AppId $AG_Item_APP"
                    $Count = $A_Count + 1
                    }
                }
            }
        If ($PackageDeploymentIDs) {
            Write-Host "Found Package:"
            Foreach ($PackageDeploymentID in $PackageDeploymentIDs) {
                $PackageID = $PackageName = (Get-WmiObject -ComputerName $SiteServer -Class SMS_Advertisement -Namespace root/SMS/site_$SiteCode -Filter "AdvertisementID='$PackageDeploymentID'").PackageID
                $PackageName = (Get-WmiObject -ComputerName $SiteServer -Class SMS_Package -Namespace root/SMS/site_$SiteCode -Filter "PackageID='$PackageID'").Name
                $ProgramName = (Get-WmiObject -ComputerName $SiteServer -Class SMS_DeploymentInfo -Namespace root/SMS/site_$SiteCode -Filter "DeploymentID='$PackageDeploymentIDs'").TargetSubName
                $FoundMachinePkgs += $PackageName
                "Package: " + $PackageName | Out-File "$LogPath\$MachineLogName" -Append
                $Id = “{0:D3}” -f $P_Count
                $M_PackageId = “M_PackageId$Id”
                $TSEnv.Value($M_PackageId) = $PackageID + ":" + "$ProgramName"
                Write-Host "$M_PackageId $PackageName"
                $P_Count = $P_Count + 1
                }
            }

        }
    If ($FoundMachineApps.Count -eq 0) {
        $TSEnv.Value("SkipMachineApplications") = "True"
        "No Machine AppsFound - Skip App Install" | Out-File "$LogPath\$MachineLogName" -Append
        Write-Host "No Machine AppsFound - Skip App Install"
        }
    If ($FoundMachinePkgs.Count -eq 0) {
        $TSEnv.Value("SkipMachinePackages") = "True"
        "No Machine Packages Found - Skip Package Install" | Out-File "$LogPath\$MachineLogName" -Append
        Write-Host "No Machine Packages Found - Skip Package Install"
        }

    }

If ($UserApps) {Get-UserApps}
Else {
    $TSEnv.Value("SkipUserApplications") = "True"
    "Skip User Apps" | Out-File "$LogPath\$MachineLogName" -Append
    Write-Host "Skip User Apps"
        }

If ($UserApps) {Get-MachineApps}
Else {
    $TSEnv.Value("SkipMachineApplications") = "True"
    $TSEnv.Value("SkipMachinePackages") = "True"
    "Skip Machine Apps And Packages" | Out-File "$LogPath\$MachineLogName" -Append
    Write-Host "Skip Machine Apps And Packages"
        }

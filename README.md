# MEMCM_Toolbox
Collection of  MEMCM PowerShell scripts

**Get-MECMDeviceVariables.ps1**

The script queries all device variables of one or all devices and optionally exports them to a CSV file.

**Set-MECMDeviceVariable.ps1**

The script creates device variables. Either a single variable to a device or multiple variables are created by importing a file. 

**Get-MECMPackagesOnDPGroups.ps1**

The script checks all packages (normal packages, driver packages, OS images, OS installers, boot images, applications) if they have been distributed to a DP group and exports the results to a CSV file.

**Invoke-MECMDPRefresh.ps1**

The script checks the distribution status of all content and retriggers the distribution.

**Sync-MECMDirectCollectionMembershipRules.ps1**

The script queries all direct membership rules of a computer and adds the new computer to the same collections.

**Import-MECMComputerFromFolder.ps1**

The script imports computers into MECM, adds them to a collection and sets variables. Three CSV files are read for the import, respectively for importing the computers (import-computer.csv), adding them to a collection (add-collection.csv) and setting the variables (set-variable.csv). 

**Remove-DirectMembership.ps1**

This script removes the direct memberships of a specified collection (User or Device).

**Invoke-DynamicAppInstall.ps1**

The script is intended for execution within an MCM task sequence and creates a list of applications that are assigned to the primary users of the device and the device itself in order to install them dynamically in the "Install application" or "Install package" step.

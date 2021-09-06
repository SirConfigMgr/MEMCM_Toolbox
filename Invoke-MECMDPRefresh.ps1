<#
    .SYNOPSIS
    Queries content distribution status.
    

    .DESCRIPTION
    The script checks the distribution status of all content and retriggers the distribution.

    .PARAMETER LogPath
    If set, Logfile is created.  (Full path and .log extension is required.) 

    .PARAMETER SiteServer
    FQDN of the MEMCM Site Server

    .PARAMETER SiteCode
    MEMCM Site Code

    .EXAMPLE
    PS> Invoke-MECMDpRefresh -$SiteServer "memcm.domain.com" -SiteCode "AA1"
    
    .EXAMPLE
    PS> Invoke-MECMDpRefresh -LogPath "C:\DP-Refresh.log" -$SiteServer "memcm.domain.com" -SiteCode "AA1"
#>

param (
    [Parameter(Mandatory=$true)][String]$SiteServer,
    [Parameter(Mandatory=$true)][String]$SiteCode,
    [String]$LogPath
    )

Function Write-Log {

[CmdletBinding()]
Param(
    [parameter(Mandatory=$true)]
    [String]$Path,

    [parameter(Mandatory=$true)]
    [String]$Message,

    [parameter(Mandatory=$true)]
    [String]$Component,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Info", "Warning", "Error")]
    [String]$Type
    )

Switch ($Type) {
    "Info" {[int]$Type = 1}
    "Warning" {[int]$Type = 2}
    "Error" {[int]$Type = 3}
    }

$Content = "<![LOG[$Message]LOG]!>" +`
        "<time=`"$(Get-Date -Format "HH:mm:ss.ffffff")`" " +`
        "date=`"$(Get-Date -Format "M-d-yyyy")`" " +`
        "component=`"$Component`" " +`
        "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " +`
        "type=`"$Type`" " +`
        "thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " +`
        "file=`"`">"

Add-Content -Path $Path -Value $Content
}


$Failures = Get-WmiObject -Namespace root\sms\site_$SiteCode -Query "SELECT * FROM SMS_PackageStatusDistPointsSummarizer WHERE State = 3 Or State = 6 Or State = 8"
$Retry = Get-WmiObject -Namespace root\sms\site_$SiteCode -Query "SELECT * FROM SMS_PackageStatusDistPointsSummarizer WHERE State = 2 or State = 5"
$Pending = Get-WmiObject -Namespace root\sms\site_$SiteCode -Query "SELECT * FROM SMS_PackageStatusDistPointsSummarizer WHERE State = 1 or State = 4"

If ($Failures -eq $null) {
    If ($LogPath) {
        $Info = "No failed content distribution."
        Write-Log -Path $LogPath -Message ($Info | Out-String) -Component "Failed" -Type Info
        }
    Write-Host -ForegroundColor White -BackgroundColor Green "No failed content distribution."
    }

Else {
    Foreach ($F in $Failures) {
        $DistributionPoints = Get-WmiObject -Namespace root\sms\site_$SiteCode -Query "SELECT * FROM SMS_DistributionPoint WHERE SiteCode='$($F.SiteCode)' AND  PackageID='$($F.PackageID)'"
        Foreach ($DistributionPoint in $DistributionPoints) {
            If ($DistributionPoint.ServerNALPath -eq $F.ServerNALPath) {
                $DistributionPoint.RefreshNow = $true
                $DistributionPoint.put()
                $DP = $DistributionPoint.ServerNALPath
                $DP = $DP.Split("\\")
                If ($LogPath) {
                    $Info = "Failed Package: " + $F.PackageID + " on " + $DP[2]
                    Write-Log -Path $LogPath -Message ($Info | Out-String) -Component "Failed" -Type Error
                    Write-Host -ForegroundColor White -BackgroundColor Red "Failed Package: " + $F.PackageID + " on " + $DP[2]
                    }
                }
            }
        }
    }

If ($Pending -eq $null){
    If ($LogPath) {
        $Info = "No pending content distribution."
        Write-Log -Path $LogPath -Message ($Info | Out-String) -Component "Pending" -Type Info
        }
    Write-Host -ForegroundColor White -BackgroundColor Green "No pending content distribution."
    }

Else {
    Foreach ($P in $Pending) {
        $DistributionPoints = Get-WmiObject -Namespace root\sms\site_$SiteCode -Query "SELECT * FROM SMS_DistributionPoint WHERE SiteCode='$($P.SiteCode)' AND  PackageID='$($P.PackageID)'"
        Foreach ($DistributionPoint in $DistributionPoints) {
            If ($DistributionPoint.ServerNALPath -eq $P.ServerNALPath) {
                $DistributionPoint.RefreshNow = $true
                $DistributionPoint.put()
                $DP = $DistributionPoint.ServerNALPath
                $DP = $DP.Split("\\")
                If ($LogPath) {
                    $Info = "Pending Package: " + $P.PackageID + " on " + $DP[2]
                    Write-Log -Path $LogPath -Message ($Info | Out-String) -Component "Pending" -Type Error
                    Write-Host -ForegroundColor White -BackgroundColor Red "Pending Package: " + $P.PackageID + " on " + $DP[2]
                    }         
                }
            }
        }
    }


If ($Retry -eq $null){
    If ($LogPath) {
        $Info = "No retry content distribution."
        Write-Log -Path $LogPath -Message ($Info | Out-String) -Component "Retry" -Type Info
        }
    Write-Host -ForegroundColor White -BackgroundColor Green "No retry content distribution."
    }

Else {
    Foreach ($R in $Retry) {
        $DistributionPoints = Get-WmiObject -Namespace root\sms\site_$SiteCode -Query "SELECT * FROM SMS_DistributionPoint WHERE SiteCode='$($R.SiteCode)' AND  PackageID='$($R.PackageID)'"
        Foreach ($DistributionPoint in $DistributionPoints) {
            If ($DistributionPoint.ServerNALPath -eq $R.ServerNALPath) {
                $DistributionPoint.RefreshNow = $true
                $DistributionPoint.put()
                $DP = $DistributionPoint.ServerNALPath
                $DP = $DP.Split("\\")
                If ($LogPath) {
                    $Info = "Retry Package: " + $R.PackageID + " on " + $DP[2]
                    Write-Log -Path $LogPath -Message ($Info | Out-String) -Component "Retry" -Type Error
                    Write-Host -ForegroundColor White -BackgroundColor Red "Retry Package: " + $R.PackageID + " on " + $DP[2]
                    } 
                }
            }
        }
    }
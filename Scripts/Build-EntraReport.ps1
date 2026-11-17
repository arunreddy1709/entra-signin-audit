<#
 Rebuilds the user activity / licence review report (JSON + Excel)
 from the collected sign-in activity and the newest licence export.

   .\Build-EntraReport.ps1 -Mode Test
   .\Build-EntraReport.ps1 -Mode Production
#>
[CmdletBinding()]
param(
    [ValidateSet('Test', 'Production')]
    [string]$Mode = 'Test'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'EntraAudit.Common.ps1')

$Paths = Get-AuditPaths -Mode $Mode
$Result = Invoke-EntraReportBuild -Paths $Paths

Write-Host "Report month : $($Result.Month)"
Write-Host "Users        : $($Result.Users)"
Write-Host "To remove    : $($Result.RemoveCount)"
Write-Host "To review    : $($Result.ReviewCount)"
Write-Host "Licence file : $(if ($Result.LicenseFile) { $Result.LicenseFile } else { 'none uploaded yet' })"
Write-Host "Excel        : $($Result.Xlsx)"

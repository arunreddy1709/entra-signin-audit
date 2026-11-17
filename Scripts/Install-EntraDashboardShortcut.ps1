<#
 Creates a desktop shortcut that opens the Entra audit dashboard app.

   .\Install-EntraDashboardShortcut.ps1 -Mode Production
   .\Install-EntraDashboardShortcut.ps1 -Mode Test
#>
[CmdletBinding()]
param(
    [ValidateSet('Test', 'Production')]
    [string]$Mode = 'Production'
)

$ErrorActionPreference = 'Stop'

$Server = Join-Path $PSScriptRoot 'Start-EntraDashboard.ps1'
if (-not (Test-Path -LiteralPath $Server)) { throw "Not found: $Server" }

$Desktop = [Environment]::GetFolderPath('Desktop')
if ($Mode -eq 'Test') { $Name = 'Entra Audit Dashboard (TEST).lnk' } else { $Name = 'Entra Audit Dashboard.lnk' }
$LinkPath = Join-Path $Desktop $Name

$Shell = New-Object -ComObject WScript.Shell
$Link = $Shell.CreateShortcut($LinkPath)
# Windowless console host (see Register-EntraCollectorTask.ps1 for why not -WindowStyle Hidden)
$PowerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$Link.TargetPath = Join-Path $env:SystemRoot 'System32\conhost.exe'
$Link.Arguments = "--headless `"$PowerShellExe`" -NoProfile -ExecutionPolicy Bypass -File `"$Server`" -Mode $Mode"
$Link.WorkingDirectory = $PSScriptRoot
$Link.WindowStyle = 7   # minimised
$Link.IconLocation = "$env:SystemRoot\System32\imageres.dll,144"
$Link.Description = "Entra ID user activity and licence review dashboard ($Mode)"
$Link.Save()

Write-Host "Shortcut created: $LinkPath"

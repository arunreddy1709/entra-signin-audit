<#
 One-off: seeds per-user "last successful sign-in" from Entra admin
 center sign-in CSV exports (e.g. the October 2026 files), so the
 dashboard has history from before the collector started.

 Only State\user_activity.json is updated (latest date wins).
 The sign-in archive and checkpoint are NOT touched.

   .\Import-SeedActivity.ps1 -Mode Test `
       -InteractiveCsv  '..\03_Monthly_Archive\oct_26_source_files\Interactive_SignIns.csv_oct_26.csv' `
       -NonInteractiveCsv '..\03_Monthly_Archive\oct_26_source_files\NonInteractive_SignIns.csv_oct_26.csv'
#>
[CmdletBinding()]
param(
    [ValidateSet('Test', 'Production')]
    [string]$Mode = 'Test',

    [string]$InteractiveCsv,
    [string]$NonInteractiveCsv
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'EntraAudit.Common.ps1')

$Paths = Get-AuditPaths -Mode $Mode
Initialize-AuditFolders $Paths
$State = Read-ActivityState $Paths
$Before = $State.Users.Count

$Inputs = @()
if ($InteractiveCsv) { $Inputs += @{ Kind = 'Interactive'; Path = $InteractiveCsv } }
if ($NonInteractiveCsv) { $Inputs += @{ Kind = 'NonInteractive'; Path = $NonInteractiveCsv } }
if ($Inputs.Count -eq 0) { throw 'Provide -InteractiveCsv and/or -NonInteractiveCsv.' }

foreach ($Src in $Inputs) {
    if (-not (Test-Path -LiteralPath $Src.Path)) { throw "File not found: $($Src.Path)" }

    $Header = Get-Content -LiteralPath $Src.Path -TotalCount 1
    foreach ($Col in @('Date (UTC)', 'User ID', 'Username', 'Status')) {
        if ($Header -notmatch [regex]::Escape("`"$Col`"")) { throw "$($Src.Path) is missing column '$Col'." }
    }

    $Kind = $Src.Kind
    $Total = 0; $Success = 0
    $Started = Get-Date
    Write-Host "Reading $Kind export: $($Src.Path)"

    # Admin center exports repeat some column names (e.g. "Incoming token type"):
    # make header names unique, then skip the original header line.
    $Names = @(); $Used = @{}
    foreach ($Raw in (($Header.Trim().Trim('"')) -split '","')) {
        $Name = $Raw.Trim(); $N = $Name; $i = 2
        while ($Used.ContainsKey($N)) { $N = "${Name}_$i"; $i++ }
        $Used[$N] = $true; $Names += $N
    }

    Import-Csv -LiteralPath $Src.Path -Header $Names | Select-Object -Skip 1 | ForEach-Object {
        $Total++
        if ($_.Status -eq 'Success' -and $_.'User ID') {
            Update-ActivityState -State $State -UserId $_.'User ID' -Upn $_.Username -Name $_.User `
                -Kind $Kind -IsoUtc (ConvertTo-UtcIso $_.'Date (UTC)')
            $Success++
        }
        if ($Total % 50000 -eq 0) { Write-Host "  $Total rows read..." }
    }

    Write-Host ("  {0}: {1} rows, {2} successful sign-ins applied ({3:N0} s)" -f $Kind, $Total, $Success, ((Get-Date) - $Started).TotalSeconds)
}

Save-ActivityState -Paths $Paths -State $State
Write-Host "User activity state saved. Users before: $Before, after: $($State.Users.Count). History since: $(Format-LocalDate $State.HistoryStartUtc)"

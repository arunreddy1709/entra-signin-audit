<#
 Creates / updates the Task Scheduler task for the sign-in collector.
 Does NOT touch the original "Entra ID Monthly Audit" task.

 Production task "Entra ID Daily Sign-In Collector":
   - Monday-Friday at 1:00 PM
   - At sign-in (+15 min) to catch up after weekends / days away
   - Missed runs start as soon as possible; runs on battery
   - Skips if the last successful collection ended < 8 hours ago
   - Max run time 12 h; never two at once

   .\Register-EntraCollectorTask.ps1 -Mode Production
   .\Register-EntraCollectorTask.ps1 -Mode Test       (one-off test task, started immediately)
   .\Register-EntraCollectorTask.ps1 -Mode Test -Remove
#>
[CmdletBinding()]
param(
    [ValidateSet('Test', 'Production')]
    [string]$Mode = 'Production',

    [string]$At = '1:00PM',

    [switch]$Remove
)

$ErrorActionPreference = 'Stop'

if ($Mode -eq 'Test') { $TaskName = 'Entra ID Daily Sign-In Collector - TEST' }
else { $TaskName = 'Entra ID Daily Sign-In Collector' }

if ($Remove) {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed task: $TaskName"
    }
    else { Write-Host "Task not found: $TaskName" }
    return
}

$Collector = Join-Path $PSScriptRoot 'Export-EntraSignIns_Weekly.ps1'
if (-not (Test-Path -LiteralPath $Collector)) { throw "Not found: $Collector" }

$User = "$env:USERDOMAIN\$env:USERNAME"
$PowerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

# Run inside a windowless console host. Using "powershell -WindowStyle Hidden"
# directly gets the process killed (0xC000013A) on Windows 11 when Windows
# Terminal is the default console, because hiding closes the terminal tab.
$HostExe = Join-Path $env:SystemRoot 'System32\conhost.exe'

if ($Mode -eq 'Test') {
    $Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$Collector`" -Mode Test -Trigger Scheduled -TestHours 1"
    # Placeholder trigger far in the future; the test is started on demand below
    $Triggers = @(New-ScheduledTaskTrigger -Once -At (Get-Date).AddDays(30))
    $Limit = New-TimeSpan -Hours 2
}
else {
    $Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$Collector`" -Mode Production -Trigger Scheduled -SkipIfRecentHours 8"
    $Weekday = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -At $At
    $Logon = New-ScheduledTaskTrigger -AtLogOn -User $User
    $Logon.Delay = 'PT15M'
    $Triggers = @($Weekday, $Logon)
    $Limit = New-TimeSpan -Hours 12
}

$Action = New-ScheduledTaskAction -Execute $HostExe -Argument "--headless `"$PowerShellExe`" $Arguments" -WorkingDirectory $PSScriptRoot

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -RunOnlyIfNetworkAvailable `
    -ExecutionTimeLimit $Limit `
    -MultipleInstances IgnoreNew

$Principal = New-ScheduledTaskPrincipal -UserId $User -LogonType Interactive -RunLevel Limited

$Task = New-ScheduledTask -Action $Action -Trigger $Triggers -Settings $Settings -Principal $Principal `
    -Description "Collects Entra ID interactive and non-interactive sign-ins ($Mode), updates the archive, licence review report and dashboard."

Register-ScheduledTask -TaskName $TaskName -InputObject $Task -Force | Out-Null
Write-Host "Task registered: $TaskName"

$Registered = Get-ScheduledTask -TaskName $TaskName
foreach ($T in $Registered.Triggers) {
    Write-Host ("  Trigger: {0} {1} {2}" -f $T.CimClass.CimClassName, $T.StartBoundary, $(if ($T.Delay) { "delay $($T.Delay)" } else { '' }))
}
Write-Host "  Action : $($Registered.Actions[0].Execute) $($Registered.Actions[0].Arguments)"
Write-Host "  Next run: $((Get-ScheduledTaskInfo -TaskName $TaskName).NextRunTime)"

if ($Mode -eq 'Test') {
    Start-ScheduledTask -TaskName $TaskName
    Write-Host 'Test task started.'
}

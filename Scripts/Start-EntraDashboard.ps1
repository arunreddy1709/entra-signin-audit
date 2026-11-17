<#
============================================================
 ENTRA AUDIT DASHBOARD - local desktop app
 Windows PowerShell 5.1

 Starts a small web server on THIS PC ONLY (127.0.0.1) and opens
 the dashboard in an Edge app window. The server stops by itself
 a few minutes after the window is closed.

 Desktop shortcut target:
   powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden
     -File "...\Scripts\Start-EntraDashboard.ps1" -Mode Production
============================================================
#>
[CmdletBinding()]
param(
    [ValidateSet('Test', 'Production')]
    [string]$Mode = 'Production',

    [int]$Port = 0,

    # Stop the server after this many minutes without any request
    [ValidateRange(1, 240)]
    [int]$IdleMinutes = 10,

    [switch]$NoBrowser
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'EntraAudit.Common.ps1')

$Paths = Get-AuditPaths -Mode $Mode
Initialize-AuditFolders $Paths

if ($Port -eq 0) {
    if ($Mode -eq 'Test') { $Port = 8766 } else { $Port = 8765 }
}

$BaseUrl = "http://127.0.0.1:$Port/"
$ServerLog = Join-Path $Paths.Logs 'Dashboard_Server.log'
$IndexPath = Join-Path $Paths.Dashboard 'index.html'
$Token = [guid]::NewGuid().ToString('N')
$Utf8 = New-Object System.Text.UTF8Encoding($false)

function Write-ServerLog {
    param([string]$Message)
    try {
        Add-Content -LiteralPath $ServerLog -Encoding UTF8 -Value ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
    }
    catch { }
}

function Open-AppWindow {
    $EdgeCandidates = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
    )
    $Edge = $EdgeCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if ($Edge) {
        Start-Process -FilePath $Edge -ArgumentList "--app=$BaseUrl"
    }
    else {
        Start-Process $BaseUrl
    }
}

# ---- Already running? Just open the window.
try {
    $Ping = Invoke-WebRequest -Uri ($BaseUrl + 'ping') -UseBasicParsing -TimeoutSec 2
    if ($Ping.Content -eq "entra-dashboard-$Mode") {
        if (-not $NoBrowser) { Open-AppWindow }
        exit 0
    }
}
catch { }

# ------------------------------------------------------------
# Response helpers
# ------------------------------------------------------------

function Send-Bytes {
    param($Ctx, [int]$Status, [byte[]]$Bytes, [string]$ContentType, [hashtable]$Headers = @{})
    $Res = $Ctx.Response
    $Res.StatusCode = $Status
    $Res.ContentType = $ContentType
    $Res.Headers['Cache-Control'] = 'no-store'
    $Res.Headers['X-Content-Type-Options'] = 'nosniff'
    foreach ($K in $Headers.Keys) { $Res.Headers[$K] = $Headers[$K] }
    $Res.ContentLength64 = $Bytes.Length
    $Res.OutputStream.Write($Bytes, 0, $Bytes.Length)
}

function Send-Text {
    param($Ctx, [int]$Status, [string]$Text, [string]$ContentType = 'text/plain; charset=utf-8')
    Send-Bytes -Ctx $Ctx -Status $Status -Bytes $Utf8.GetBytes($Text) -ContentType $ContentType
}

function Send-Json {
    param($Ctx, [int]$Status, $Object)
    if ($Object -is [string]) { $Json = $Object } else { $Json = ConvertTo-Json -InputObject $Object -Depth 6 -Compress }
    Send-Text -Ctx $Ctx -Status $Status -Text $Json -ContentType 'application/json; charset=utf-8'
}

function Read-Body {
    param($Ctx, [int]$MaxBytes = 20MB)
    if ($Ctx.Request.ContentLength64 -gt $MaxBytes) { throw "Upload too large (max $([int]($MaxBytes / 1MB)) MB)." }
    $Reader = New-Object System.IO.StreamReader($Ctx.Request.InputStream, [System.Text.Encoding]::UTF8)
    try { return $Reader.ReadToEnd() } finally { $Reader.Dispose() }
}

# ------------------------------------------------------------
# State helpers
# ------------------------------------------------------------

# Graph pages can take ~2 min, so give the graceful stop time before forcing
$ForceStopAfterSeconds = 240

function Get-RunState {
    $LastRun = Read-JsonFile $Paths.LastRun
    $Running = $false
    $Proc = $null
    if ($null -ne $LastRun -and $LastRun.Status -eq 'Running' -and $LastRun.Pid) {
        $Proc = Get-Process -Id ([int]$LastRun.Pid) -ErrorAction SilentlyContinue
        if ($null -ne $Proc -and $Proc.ProcessName -match 'powershell') { $Running = $true }
    }

    $StopRequested = Test-Path -LiteralPath $Paths.StopRequest

    # Graceful stop did not happen in time -> force-stop the collector process
    if ($Running -and $StopRequested) {
        $Age = ((Get-Date) - (Get-Item -LiteralPath $Paths.StopRequest).LastWriteTime).TotalSeconds
        if ($Age -gt $ForceStopAfterSeconds) {
            try { Stop-Process -Id $Proc.Id -Force -ErrorAction Stop } catch { }
            Start-Sleep -Milliseconds 500
            Remove-Item -LiteralPath $Paths.StopRequest -Force -ErrorAction SilentlyContinue
            $Msg = 'Force-stopped from the dashboard (did not stop by itself within 4 minutes). Days already saved are kept; the next run continues from there.'
            Write-JsonFile -Path $Paths.LastRun -Object ([ordered]@{
                    Status = 'Stopped'; RunId = $LastRun.RunId; Trigger = $LastRun.Trigger; Mode = $Mode
                    StartedLocal = $LastRun.StartedLocal; FinishedLocal = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    Message = $Msg; LogFile = $LastRun.LogFile
                })
            Add-AuditNotification -Paths $Paths -Level 'warning' -Title "Entra sign-in collection stopped ($Mode)" -Message $Msg
            Write-ServerLog "Collector PID $($Proc.Id) force-stopped."
            $LastRun = Read-JsonFile $Paths.LastRun
            $Running = $false
            $StopRequested = $false
        }
    }

    [PSCustomObject]@{ LastRun = $LastRun; Running = $Running; StopRequested = ($Running -and $StopRequested) }
}

$script:ScheduleCache = $null
$script:ScheduleCacheTime = [datetime]::MinValue

function Get-ScheduleState {
    param([switch]$Fresh)
    if (-not $Fresh -and ((Get-Date) - $script:ScheduleCacheTime).TotalSeconds -lt 60 -and $null -ne $script:ScheduleCache) {
        return $script:ScheduleCache
    }

    $Task = Get-ScheduledTask -TaskName $Paths.TaskName -ErrorAction SilentlyContinue
    if ($null -eq $Task) {
        $State = [ordered]@{ exists = $false; enabled = $false; nextRun = ''; taskName = $Paths.TaskName }
    }
    else {
        $Info = Get-ScheduledTaskInfo -TaskName $Paths.TaskName -ErrorAction SilentlyContinue
        $Next = ''
        if ($Info -and $Info.NextRunTime -and $Info.NextRunTime.Year -gt 2000) { $Next = $Info.NextRunTime.ToString('yyyy-MM-dd HH:mm') }
        $State = [ordered]@{ exists = $true; enabled = ($Task.State -ne 'Disabled'); nextRun = $Next; taskName = $Paths.TaskName }
    }

    $script:ScheduleCache = $State
    $script:ScheduleCacheTime = Get-Date
    return $State
}

function Get-LogTail {
    param([int]$Lines = 15)
    $Latest = Get-ChildItem -LiteralPath $Paths.Logs -Filter 'Collector_*.log' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $Latest) { return @() }
    return @(Get-Content -LiteralPath $Latest.FullName -Tail $Lines -Encoding UTF8)
}

function Get-MonthList {
    if (-not (Test-Path -LiteralPath $Paths.Reports)) { return @() }
    return @(Get-ChildItem -LiteralPath $Paths.Reports -Directory |
            Where-Object { $_.Name -match '^\d{4}-\d{2}$' -and (Test-Path -LiteralPath (Join-Path $_.FullName 'report.json')) } |
            Sort-Object Name -Descending | ForEach-Object { $_.Name })
}

function Get-MonthParam {
    param($Ctx)
    $M = [string]$Ctx.Request.QueryString['month']
    if ([string]::IsNullOrEmpty($M)) { return '' }
    if ($M -notmatch '^\d{4}-\d{2}$') { throw 'Invalid month.' }
    return $M
}

# ------------------------------------------------------------
# Routes
# ------------------------------------------------------------

function Invoke-Route {
    param($Ctx)

    $Req = $Ctx.Request
    $Path = $Req.Url.AbsolutePath
    $Method = $Req.HttpMethod

    # Only answer requests addressed to this PC (blocks DNS-rebinding tricks)
    $HostHeader = [string]$Req.Headers['Host']
    if ($HostHeader -ne "127.0.0.1:$Port" -and $HostHeader -ne "localhost:$Port") {
        Send-Text $Ctx 403 'Forbidden'
        return
    }

    if ($Path -eq '/ping') {
        Send-Text $Ctx 200 "entra-dashboard-$Mode"
        return
    }

    if ($Path -eq '/' -or $Path -eq '/index.html') {
        $Html = [System.IO.File]::ReadAllText($IndexPath)
        $Html = $Html.Replace('__AUTH_TOKEN__', $Token).Replace('__MODE__', $Mode)
        Send-Text $Ctx 200 $Html 'text/html; charset=utf-8'
        return
    }

    if (-not $Path.StartsWith('/api/')) {
        Send-Text $Ctx 404 'Not found'
        return
    }

    # Every API call needs this session's token
    $Supplied = [string]$Req.Headers['X-Auth-Token']
    if ([string]::IsNullOrEmpty($Supplied)) { $Supplied = [string]$Req.QueryString['t'] }
    if ($Supplied -ne $Token) {
        Send-Json $Ctx 401 @{ error = 'Session expired. Close and reopen the dashboard.' }
        return
    }

    switch ("$Method $Path") {

        'GET /api/report' {
            $Month = Get-MonthParam $Ctx
            if ($Month) { $File = Join-Path (Join-Path $Paths.Reports $Month) 'report.json' }
            else { $File = Join-Path $Paths.Reports 'latest.json' }

            if (Test-Path -LiteralPath $File) {
                Send-Json $Ctx 200 ([System.IO.File]::ReadAllText($File))
            }
            else {
                Send-Json $Ctx 200 @{ empty = $true }
            }
        }

        'GET /api/months' {
            Send-Json $Ctx 200 @{ months = @(Get-MonthList) }
        }

        'GET /api/status' {
            $State = Get-RunState
            $Notes = Read-JsonFile $Paths.Notifications
            $Checkpoint = Read-JsonFile $Paths.Checkpoint
            $License = Get-LatestLicenseFile $Paths
            Send-Json $Ctx 200 ([ordered]@{
                    mode          = $Mode
                    running       = $State.Running
                    stopRequested = $State.StopRequested
                    schedule      = (Get-ScheduleState)
                    lastRun       = $State.LastRun
                    checkpoint    = $Checkpoint
                    licenseFile   = $(if ($License) { $License.Name } else { '' })
                    notifications = @($Notes | Select-Object -First 15)
                    logTail       = @(Get-LogTail)
                    nowLocal      = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                })
        }

        'POST /api/run' {
            $State = Get-RunState
            if ($State.Running) {
                Send-Json $Ctx 409 @{ error = 'A collection is already running.' }
                break
            }

            $ArgList = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$($Paths.Collector)`" -Mode $Mode -Trigger Manual"
            Start-Process -FilePath 'powershell.exe' -ArgumentList $ArgList -WindowStyle Hidden | Out-Null
            Write-ServerLog "Manual collection started from dashboard ($Mode)."
            Send-Json $Ctx 202 @{ started = $true }
        }

        'POST /api/stop' {
            $State = Get-RunState
            if (-not $State.Running) {
                Send-Json $Ctx 409 @{ error = 'No collection is running.' }
                break
            }
            Set-Content -LiteralPath $Paths.StopRequest -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -Encoding UTF8
            Write-ServerLog "Stop requested from dashboard for collector PID $($State.LastRun.Pid)."
            Send-Json $Ctx 202 @{ stopping = $true }
        }

        'POST /api/schedule' {
            $Body = Read-Body $Ctx -MaxBytes 1KB
            $Wanted = [bool](($Body | ConvertFrom-Json).enabled)
            $Task = Get-ScheduledTask -TaskName $Paths.TaskName -ErrorAction SilentlyContinue
            if ($null -eq $Task) {
                Send-Json $Ctx 404 @{ error = "Scheduled task '$($Paths.TaskName)' is not set up yet." }
                break
            }
            if ($Wanted) { Enable-ScheduledTask -TaskName $Paths.TaskName | Out-Null }
            else { Disable-ScheduledTask -TaskName $Paths.TaskName | Out-Null }

            $Word = $(if ($Wanted) { 'resumed' } else { 'paused' })
            Write-ServerLog "Automatic updates $Word from dashboard."
            Add-AuditNotification -Paths $Paths -Level 'info' -Title "Automatic updates $Word" `
                -Message $(if ($Wanted) { 'Weekday 1 PM and sign-in catch-up runs are on again.' } else { 'No scheduled collections until resumed. Update now still works.' })
            Send-Json $Ctx 200 @{ ok = $true; schedule = (Get-ScheduleState -Fresh) }
        }

        'POST /api/upload-licenses' {
            $Body = Read-Body $Ctx
            $Body = $Body.TrimStart([char]0xFEFF)
            $FirstLine = ($Body -split "`r?`n", 2)[0]
            $Problem = Test-LicenseCsvHeader $FirstLine
            if ($Problem) {
                Send-Json $Ctx 400 @{ error = $Problem }
                break
            }

            $Original = [string]$Req.Headers['X-File-Name']
            $Safe = ([System.IO.Path]::GetFileNameWithoutExtension($Original) -replace '[^A-Za-z0-9_\-]', '_')
            if ($Safe.Length -gt 60) { $Safe = $Safe.Substring(0, 60) }
            $Target = Join-Path $Paths.Licenses ("Licenses_{0}_{1}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'), $Safe)

            [System.IO.File]::WriteAllText($Target, $Body, (New-Object System.Text.UTF8Encoding($true)))
            Write-ServerLog "Licence export uploaded: $Target"

            $Result = Invoke-EntraReportBuild -Paths $Paths
            Add-AuditNotification -Paths $Paths -Level 'info' -Title 'Licence export uploaded' `
                -Message "$Original - report rebuilt: $($Result.RemoveCount) to remove, $($Result.ReviewCount) to review."
            Send-Json $Ctx 200 @{ ok = $true; saved = (Split-Path -Leaf $Target); result = $Result }
        }

        'POST /api/rebuild' {
            $Result = Invoke-EntraReportBuild -Paths $Paths
            Send-Json $Ctx 200 @{ ok = $true; result = $Result }
        }

        'GET /api/export' {
            $Month = Get-MonthParam $Ctx
            if ($Month) {
                $File = Join-Path (Join-Path $Paths.Reports $Month) "Entra_User_Activity_$Month.xlsx"
                $Name = "Entra_User_Activity_$Month.xlsx"
            }
            else {
                $File = Join-Path $Paths.Reports 'latest.xlsx'
                $Name = "Entra_User_Activity_$((Get-Date).ToString('yyyy-MM-dd')).xlsx"
            }

            if (-not (Test-Path -LiteralPath $File)) {
                Send-Text $Ctx 404 'No Excel report yet. Upload a licence export or run an update first.'
                break
            }

            Send-Bytes -Ctx $Ctx -Status 200 -Bytes ([System.IO.File]::ReadAllBytes($File)) `
                -ContentType 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' `
                -Headers @{ 'Content-Disposition' = "attachment; filename=`"$Name`"" }
        }

        'POST /api/open-folder' {
            Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$($Paths.Reports)`""
            Send-Json $Ctx 200 @{ ok = $true }
        }

        default {
            Send-Json $Ctx 404 @{ error = "Unknown endpoint $Method $Path" }
        }
    }
}

# ------------------------------------------------------------
# Server loop
# ------------------------------------------------------------

$Listener = New-Object System.Net.HttpListener
$Listener.Prefixes.Add($BaseUrl)

try {
    $Listener.Start()
}
catch {
    if ($_.Exception.Message -match 'conflicts with an existing registration') {
        # A dashboard server for this mode is already running - just show the window
        Write-ServerLog 'Server already running; opening window only.'
        if (-not $NoBrowser) { Open-AppWindow }
        exit 0
    }
    Write-ServerLog "Could not start listener on $BaseUrl : $($_.Exception.Message)"
    throw
}

Write-ServerLog "Dashboard server started on $BaseUrl (mode $Mode)."
Write-Host "Entra dashboard running at $BaseUrl (mode $Mode). Stops after $IdleMinutes idle minutes."

if (-not $NoBrowser) { Open-AppWindow }

$LastRequest = Get-Date

try {
    while ($Listener.IsListening) {
        $Pending = $Listener.GetContextAsync()

        while (-not $Pending.AsyncWaitHandle.WaitOne(1000)) {
            if (((Get-Date) - $LastRequest).TotalMinutes -ge $IdleMinutes) {
                Write-ServerLog 'No requests for a while - dashboard server stopping.'
                return
            }
        }

        $Ctx = $Pending.GetAwaiter().GetResult()
        $LastRequest = Get-Date

        try {
            Invoke-Route $Ctx
        }
        catch {
            Write-ServerLog "Error on $($Ctx.Request.HttpMethod) $($Ctx.Request.Url.AbsolutePath): $($_.Exception.Message)"
            try { Send-Json $Ctx 500 @{ error = $_.Exception.Message } } catch { }
        }
        finally {
            try { $Ctx.Response.Close() } catch { }
        }
    }
}
finally {
    try { $Listener.Stop(); $Listener.Close() } catch { }
}

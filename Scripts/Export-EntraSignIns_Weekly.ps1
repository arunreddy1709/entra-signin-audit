<#
============================================================
 ENTRA ID SIGN-IN COLLECTOR  (scheduled DAILY; name kept for history)
 Windows PowerShell 5.1 - Microsoft Graph AppOnly

 - Window: last successful checkpoint minus overlap -> now (UTC)
 - Small time slices; a slice that hits HTTP 410 / timeout /
   exhausted retries is split in half and re-queried (expired
   nextLinks are never reused)
 - Both streams are fully staged before anything is committed
 - Archive: DAILY CSV partitions (local date) in month folders,
   deduplicated by sign-in Id
 - Per-user last successful sign-in kept in State\user_activity.json
 - Checkpoint is written LAST, only after both streams commit
 - After the run: report rebuilt, desktop notification shown

 Modes:
   -Mode Test        writes only to 06_Scheduler_Test\Weekly_Test
   -Mode Production  writes to 07_Weekly_Collector

 Examples:
   .\Export-EntraSignIns_Weekly.ps1 -Mode Test -TestHours 6
   .\Export-EntraSignIns_Weekly.ps1 -Mode Test -SimulateFailure BeforeCheckpoint
   .\Export-EntraSignIns_Weekly.ps1 -Mode Production
============================================================
#>

[CmdletBinding()]
param(
    [ValidateSet('Test', 'Production')]
    [string]$Mode = 'Test',

    # Who started this run (shown in logs, notifications and dashboard)
    [ValidateSet('Scheduled', 'Manual')]
    [string]$Trigger = 'Scheduled',

    # Test mode: maximum look-back in hours
    [ValidateRange(1, 168)]
    [int]$TestHours = 6,

    # Re-read this many hours before the last checkpoint to catch late records
    [ValidateRange(0, 168)]
    [int]$OverlapHours = 6,

    [ValidateRange(1, 24)]
    [int]$SliceHours = 6,

    # The window is committed + checkpointed in chunks of this size, so an
    # interrupted catch-up run (laptop closed) keeps the days already done
    [ValidateRange(1, 168)]
    [int]$ChunkHours = 24,

    # Exit quietly if the last successful collection ended less than this
    # many hours ago (used by the at-logon catch-up trigger). 0 = always run.
    [ValidateRange(0, 72)]
    [int]$SkipIfRecentHours = 0,

    [ValidateRange(50, 999)]
    [int]$PageSize = 500,

    # A failing slice is split until it is this small, then the run fails
    [ValidateRange(1, 60)]
    [int]$MinSliceMinutes = 15,

    # Production: look-back on the very first run (no checkpoint yet)
    [ValidateRange(1, 29)]
    [int]$FirstRunDays = 3,

    # Production: never query further back than this (P2 retention = 30 days)
    [ValidateRange(1, 30)]
    [int]$MaxLookbackDays = 29,

    # Test mode only: deliberately fail to prove recovery behaviour
    [ValidateSet('None', 'AfterInteractive', 'BeforeCheckpoint')]
    [string]$SimulateFailure = 'None',

    # Skip the toast notification (report is still rebuilt)
    [switch]$NoNotification
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'EntraAudit.Common.ps1')

# ============================================================
# 1. CONFIGURATION
# ============================================================

# Tenant and app registration come from config.json in the project root
# (copy config.example.json; config.json is git-ignored). The client secret
# is never stored in config: it is a DPAPI-encrypted CLIXML file that only
# this Windows user on this PC can decrypt (see README "Store the secret").
$ConfigPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'config.json'
if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Missing $ConfigPath. Copy config.example.json to config.json and fill in TenantId and ClientId."
}
$Config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$TenantId = [string]$Config.TenantId
$ClientId = [string]$Config.ClientId
if (-not $TenantId -or -not $ClientId -or $TenantId -like '<*' -or $ClientId -like '<*') {
    throw "config.json must contain your TenantId and ClientId."
}
$SecretFolder = 'EntraMonthlyAudit'
if ($Config.SecretFolder) { $SecretFolder = [string]$Config.SecretFolder }
$SecretFilePath = Join-Path $env:APPDATA "$SecretFolder\ClientSecret.xml"

$Paths = Get-AuditPaths -Mode $Mode
Initialize-AuditFolders $Paths

$Streams = @(
    @{
        Name    = 'Interactive'
        BaseUri = 'https://graph.microsoft.com/v1.0/auditLogs/signIns'
        Filter  = 'isInteractive eq true'
    },
    @{
        Name    = 'NonInteractive'
        BaseUri = 'https://graph.microsoft.com/beta/auditLogs/signIns'
        Filter  = "signInEventTypes/any(t:t eq 'nonInteractiveUser')"
    }
)

foreach ($Stream in $Streams) {
    $StreamArchive = Join-Path $Paths.Archive $Stream.Name
    if (-not (Test-Path -LiteralPath $StreamArchive)) {
        New-Item -ItemType Directory -Path $StreamArchive -Force | Out-Null
    }
}

$IsoFormat = $script:EntraIsoFormat
$Invariant = $script:EntraInvariant

# ============================================================
# 2. LOGGING
# ============================================================

$RunId = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile = Join-Path $Paths.Logs "Collector_$RunId.log"
$RunStaging = Join-Path $Paths.Staging $RunId

function Write-AuditLog {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )

    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogEntry = "[$Timestamp] [$Level] $Message"

    Write-Host $LogEntry
    Add-Content -LiteralPath $LogFile -Value $LogEntry -Encoding UTF8
}

# ============================================================
# 3. GRAPH HELPERS
# ============================================================

function Get-Field {
    param($Hash, [string]$Key)
    if ($null -ne $Hash) { return $Hash[$Key] }
    return $null
}

function ConvertTo-CsvRow {
    param($Record)

    $Status = Get-Field $Record 'status'
    $Device = Get-Field $Record 'deviceDetail'
    $Location = Get-Field $Record 'location'

    [PSCustomObject]@{
        Id                      = $Record['id']
        CreatedDateTime         = ConvertTo-UtcIso $Record['createdDateTime']
        UserDisplayName         = $Record['userDisplayName']
        UserPrincipalName       = $Record['userPrincipalName']
        UserId                  = $Record['userId']
        AppDisplayName          = $Record['appDisplayName']
        ClientAppUsed           = $Record['clientAppUsed']
        IPAddress               = $Record['ipAddress']
        IsInteractive           = $Record['isInteractive']
        ConditionalAccessStatus = $Record['conditionalAccessStatus']
        ErrorCode               = Get-Field $Status 'errorCode'
        FailureReason           = Get-Field $Status 'failureReason'
        OperatingSystem         = Get-Field $Device 'operatingSystem'
        Browser                 = Get-Field $Device 'browser'
        City                    = Get-Field $Location 'city'
        CountryOrRegion         = Get-Field $Location 'countryOrRegion'
    }
}

function Get-GraphErrorInfo {
    param($ErrorRecord)

    $StatusCode = 0
    $RetryAfter = 0
    $IsTimeout = $false

    $Ex = $ErrorRecord.Exception
    while ($null -ne $Ex) {
        if ($Ex -is [System.Threading.Tasks.TaskCanceledException] -or
            $Ex -is [System.TimeoutException] -or
            $Ex -is [System.OperationCanceledException]) {
            $IsTimeout = $true
        }

        if ($StatusCode -eq 0 -and $Ex.PSObject.Properties['Response'] -and $null -ne $Ex.Response) {
            try { $StatusCode = [int]$Ex.Response.StatusCode } catch { }
            try {
                $Delta = $Ex.Response.Headers.RetryAfter.Delta
                if ($Delta) { $RetryAfter = [int][Math]::Ceiling($Delta.TotalSeconds) }
            }
            catch { }
        }

        $Ex = $Ex.InnerException
    }

    $Message = [string]$ErrorRecord.Exception.Message

    if ($StatusCode -eq 0) {
        if ($Message -match '\bGone\b') { $StatusCode = 410 }
        elseif ($Message -match 'TooManyRequests') { $StatusCode = 429 }
        elseif ($Message -match 'ServiceUnavailable|Service Unavailable') { $StatusCode = 503 }
        elseif ($Message -match 'GatewayTimeout|Gateway Timeout') { $StatusCode = 504 }
        elseif ($Message -match 'BadGateway|Bad Gateway') { $StatusCode = 502 }
        elseif ($Message -match 'InternalServerError|Internal Server Error') { $StatusCode = 500 }
    }

    if ($Message -match 'task was canceled|operation was canceled|timed out|timeout') {
        $IsTimeout = $true
    }

    # Retry the same request in place
    $Retryable = $IsTimeout -or ($StatusCode -in @(429, 500, 502, 503, 504))

    # Give up on this request and re-query the time range in smaller pieces
    $Splittable = $Retryable -or ($StatusCode -eq 410)

    if ($IsTimeout) { $Description = "timeout ($Message)" }
    elseif ($StatusCode -gt 0) { $Description = "HTTP $StatusCode ($Message)" }
    else { $Description = $Message }

    [PSCustomObject]@{
        StatusCode  = $StatusCode
        IsTimeout   = $IsTimeout
        Retryable   = $Retryable
        Splittable  = $Splittable
        RetryAfter  = $RetryAfter
        Description = $Description
    }
}

$script:RetryCount = 0
$script:SplitCount = 0
$script:PageCount = 0

$StopMessage = 'Stopped by user from the dashboard.'

function Assert-NotStopRequested {
    # The dashboard's "Stop update" button creates State\stop.request
    if (Test-Path -LiteralPath $Paths.StopRequest) {
        throw $StopMessage
    }
}

function Invoke-GraphWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [string]$Context = ''
    )

    $MaxAttempts = 5

    for ($Attempt = 1; $Attempt -le $MaxAttempts; $Attempt++) {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType HashTable -ErrorAction Stop
        }
        catch {
            $Info = Get-GraphErrorInfo $_

            if (-not $Info.Retryable -or $Attempt -eq $MaxAttempts) {
                throw
            }

            $Delay = [int][Math]::Min(60, 5 * [Math]::Pow(2, $Attempt - 1))
            if ($Info.RetryAfter -gt $Delay) { $Delay = [Math]::Min(120, $Info.RetryAfter) }

            $script:RetryCount++
            Write-AuditLog "$Context : $($Info.Description). Retry $Attempt of $($MaxAttempts - 1) in $Delay seconds." 'WARNING'
            Start-Sleep -Seconds $Delay
        }
    }
}

function New-SliceUri {
    param($Stream, [datetime]$Start, [datetime]$End)

    $S = $Start.ToString($IsoFormat, $Invariant)
    $E = $End.ToString($IsoFormat, $Invariant)

    return "$($Stream.BaseUri)?`$filter=createdDateTime ge $S and createdDateTime lt $E and $($Stream.Filter)&`$top=$PageSize"
}

function Get-SliceRows {
    # Returns all rows for [Start, End). On a splittable failure the
    # partial rows are discarded and each half is queried fresh.
    param($Stream, [datetime]$Start, [datetime]$End)

    $Label = "$($Stream.Name) $($Start.ToString($IsoFormat, $Invariant)) -> $($End.ToString($IsoFormat, $Invariant))"
    $Rows = New-Object 'System.Collections.Generic.List[object]'
    $Uri = New-SliceUri -Stream $Stream -Start $Start -End $End
    $Page = 0

    try {
        while ($Uri) {
            $Page++
            Assert-NotStopRequested
            $Response = Invoke-GraphWithRetry -Uri $Uri -Context "$Label page $Page"
            $script:PageCount++

            foreach ($Record in @($Response['value'])) {
                if ($null -ne $Record) {
                    $Rows.Add((ConvertTo-CsvRow $Record))
                }
            }

            $Uri = $Response['@odata.nextLink']

            if ($Uri) {
                Write-AuditLog "$Label page $Page done; rows so far in slice: $($Rows.Count)."
            }
        }

        return , $Rows
    }
    catch {
        if ($_.Exception.Message -eq $StopMessage) { throw }

        $Info = Get-GraphErrorInfo $_
        $SpanMinutes = ($End - $Start).TotalMinutes

        if ($Info.Splittable -and $SpanMinutes -ge (2 * $MinSliceMinutes)) {
            $script:SplitCount++
            $Mid = $Start.AddMinutes([Math]::Floor($SpanMinutes / 2))

            Write-AuditLog "$Label failed on page $Page : $($Info.Description). Discarding $($Rows.Count) partial rows; re-querying as two halves." 'WARNING'

            $Left = Get-SliceRows -Stream $Stream -Start $Start -End $Mid
            $Right = Get-SliceRows -Stream $Stream -Start $Mid -End $End

            $Combined = New-Object 'System.Collections.Generic.List[object]'
            $Combined.AddRange($Left)
            $Combined.AddRange($Right)
            return , $Combined
        }

        Write-AuditLog "$Label failed on page $Page and cannot be split further: $($Info.Description)" 'ERROR'
        throw
    }
}

# ============================================================
# 4. STAGING AND ARCHIVE
# ============================================================

function Save-StreamToStaging {
    param($Stream, [datetime]$WindowStart, [datetime]$WindowEnd)

    $StreamStaging = Join-Path $script:ChunkStaging $Stream.Name
    New-Item -ItemType Directory -Path $StreamStaging -Force | Out-Null

    Write-AuditLog "$($Stream.Name) collection started ($SliceHours-hour slices, page size $PageSize)."

    $Total = 0
    $SliceStart = $WindowStart

    while ($SliceStart -lt $WindowEnd) {
        $SliceEnd = $SliceStart.AddHours($SliceHours)
        if ($SliceEnd -gt $WindowEnd) { $SliceEnd = $WindowEnd }

        $Rows = Get-SliceRows -Stream $Stream -Start $SliceStart -End $SliceEnd

        $FileName = '{0}_{1}.csv' -f $SliceStart.ToString('yyyyMMddHHmm', $Invariant), $SliceEnd.ToString('yyyyMMddHHmm', $Invariant)
        $SlicePath = Join-Path $StreamStaging $FileName

        if ($Rows.Count -gt 0) {
            $Partial = "$SlicePath.partial"
            $Rows | Export-Csv -LiteralPath $Partial -NoTypeInformation -Encoding UTF8
            Move-Item -LiteralPath $Partial -Destination $SlicePath -Force
        }

        $Total += $Rows.Count
        Write-AuditLog "$($Stream.Name) slice $($SliceStart.ToString($IsoFormat, $Invariant)) -> $($SliceEnd.ToString($IsoFormat, $Invariant)): $($Rows.Count) rows (stream total $Total)."

        $SliceStart = $SliceEnd
    }

    Write-AuditLog "$($Stream.Name) collection completed and staged. Rows: $Total"
    return $Total
}

function Publish-StreamToArchive {
    # Commits staged rows to daily partitions (dedupe by Id) and updates
    # the per-user activity state with successful sign-ins.
    param([string]$StreamName, $ActivityState)

    $StreamStaging = Join-Path $script:ChunkStaging $StreamName
    $StreamArchive = Join-Path $Paths.Archive $StreamName

    $StagedRows = @()
    if (Test-Path -LiteralPath $StreamStaging) {
        $StagedFiles = @(Get-ChildItem -LiteralPath $StreamStaging -Filter '*.csv' -File)
        $StagedRows = @(foreach ($File in $StagedFiles) { Import-Csv -LiteralPath $File.FullName })
    }

    Write-AuditLog "$StreamName commit started. Staged rows: $($StagedRows.Count)"

    # Group by daily partition (local date)
    $ByDay = @{}
    foreach ($Row in $StagedRows) {
        $Day = (ConvertTo-UtcDate $Row.CreatedDateTime).ToLocalTime().ToString('yyyy-MM-dd', $Invariant)
        if (-not $ByDay.ContainsKey($Day)) {
            $ByDay[$Day] = New-Object 'System.Collections.Generic.List[object]'
        }
        $ByDay[$Day].Add($Row)
    }

    $TotalNew = 0
    $TotalSkipped = 0

    foreach ($Day in ($ByDay.Keys | Sort-Object)) {
        $MonthFolder = Join-Path $StreamArchive $Day.Substring(0, 7)
        if (-not (Test-Path -LiteralPath $MonthFolder)) {
            New-Item -ItemType Directory -Path $MonthFolder -Force | Out-Null
        }
        $FinalPath = Join-Path $MonthFolder "${StreamName}_SignIns_$Day.csv"

        $KnownIds = New-Object 'System.Collections.Generic.HashSet[string]'
        $ExistingCount = 0

        if (Test-Path -LiteralPath $FinalPath) {
            foreach ($Existing in (Import-Csv -LiteralPath $FinalPath)) {
                [void]$KnownIds.Add([string]$Existing.Id)
                $ExistingCount++
            }
        }

        $NewRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($Row in $ByDay[$Day]) {
            if ([string]::IsNullOrEmpty($Row.Id)) {
                throw "$StreamName staged row without Id (partition $Day)."
            }
            if ($KnownIds.Add([string]$Row.Id)) {
                $NewRows.Add($Row)
            }
        }

        $Skipped = $ByDay[$Day].Count - $NewRows.Count
        $TotalSkipped += $Skipped

        if ($NewRows.Count -eq 0) {
            Write-AuditLog "$StreamName $Day : 0 new rows ($Skipped duplicates skipped). Partition unchanged."
            continue
        }

        $TempPath = "$FinalPath.new_$RunId"

        if (Test-Path -LiteralPath $FinalPath) {
            Copy-Item -LiteralPath $FinalPath -Destination $TempPath -Force
            $NewRows | Export-Csv -LiteralPath $TempPath -NoTypeInformation -Encoding UTF8 -Append
        }
        else {
            $NewRows | Export-Csv -LiteralPath $TempPath -NoTypeInformation -Encoding UTF8
        }

        # Validate before publishing
        $ExpectedCount = $ExistingCount + $NewRows.Count
        $CheckIds = New-Object 'System.Collections.Generic.HashSet[string]'
        $CheckCount = 0
        foreach ($Check in (Import-Csv -LiteralPath $TempPath)) {
            [void]$CheckIds.Add([string]$Check.Id)
            $CheckCount++
        }

        if ($CheckCount -ne $ExpectedCount -or $CheckIds.Count -ne $ExpectedCount) {
            throw "$StreamName $Day validation failed: expected $ExpectedCount rows, found $CheckCount rows / $($CheckIds.Count) unique Ids. Temp file kept: $TempPath"
        }

        Publish-File -Source $TempPath -Destination $FinalPath
        if (Test-Path -LiteralPath "$FinalPath.bak") { Remove-Item -LiteralPath "$FinalPath.bak" -Force }

        $TotalNew += $NewRows.Count
        Write-AuditLog "$StreamName $Day : +$($NewRows.Count) new rows, $Skipped duplicates skipped, partition now $ExpectedCount rows (validated, all Ids unique)."
    }

    # Per-user last successful sign-in (idempotent: keeps the max)
    $SuccessCount = 0
    foreach ($Row in $StagedRows) {
        if ([string]$Row.ErrorCode -eq '0') {
            Update-ActivityState -State $ActivityState -UserId $Row.UserId -Upn $Row.UserPrincipalName `
                -Name $Row.UserDisplayName -Kind $StreamName -IsoUtc $Row.CreatedDateTime
            $SuccessCount++
        }
    }

    Write-AuditLog "$StreamName commit completed. New: $TotalNew, duplicates skipped: $TotalSkipped, successful sign-ins applied to user activity: $SuccessCount"

    [PSCustomObject]@{
        Staged  = $StagedRows.Count
        New     = $TotalNew
        Skipped = $TotalSkipped
    }
}

# ============================================================
# 5. MAIN EXECUTION
# ============================================================

# ---- Catch-up trigger: nothing to do if a collection finished recently
if ($SkipIfRecentHours -gt 0 -and (Test-Path -LiteralPath $Paths.Checkpoint)) {
    try {
        $RecentCheckpoint = Read-JsonFile $Paths.Checkpoint
        $AgeHours = ((Get-Date).ToUniversalTime() - (ConvertTo-UtcDate $RecentCheckpoint.LastSuccessfulEndUtc)).TotalHours
        if ($AgeHours -lt $SkipIfRecentHours) {
            Write-Host ("Last successful collection ended {0:N1} h ago (less than {1} h). Nothing to do." -f $AgeHours, $SkipIfRecentHours)
            exit 0
        }
    }
    catch { }
}

$script:ChunkStaging = $RunStaging

$GraphConnected = $false
$Succeeded = $false
$LockStream = $null
$FailureMessage = ''
$WasStopped = $false
$Committed = @{}
$StartedAt = Get-Date

Write-AuditLog '============================================'
Write-AuditLog 'ENTRA ID SIGN-IN COLLECTOR STARTED'
Write-AuditLog "Run ID: $RunId"
Write-AuditLog "Mode: $Mode   Trigger: $Trigger"
Write-AuditLog "Data folder: $($Paths.DataRoot)"
Write-AuditLog '============================================'

# ---- Single-instance lock (released automatically if the process dies).
#      If another run holds it, this run is simply skipped - not a failure.
try {
    $LockStream = New-Object System.IO.FileStream(
        $Paths.Lock,
        [System.IO.FileMode]::OpenOrCreate,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None,
        1,
        [System.IO.FileOptions]::DeleteOnClose
    )
}
catch {
    Write-AuditLog 'Another collector run is already active - this run is skipped (nothing changed).' 'WARNING'
    Write-AuditLog 'FINAL STATUS: SKIPPED'
    exit 0
}

try {
    if ($SimulateFailure -ne 'None' -and $Mode -ne 'Test') {
        throw '-SimulateFailure is only allowed in Test mode.'
    }

    Set-KeepAwake $true

    # A stop request left over from an earlier run must not stop this one
    if (Test-Path -LiteralPath $Paths.StopRequest) { Remove-Item -LiteralPath $Paths.StopRequest -Force }

    Write-JsonFile -Path $Paths.LastRun -Object ([ordered]@{
            Status = 'Running'; RunId = $RunId; Trigger = $Trigger; Mode = $Mode; Pid = $PID
            StartedLocal = $StartedAt.ToString('yyyy-MM-dd HH:mm:ss', $Invariant); FinishedLocal = ''
            Message = 'Collection in progress'; LogFile = $LogFile
        })

    # ---- Checkpoint
    $Checkpoint = $null
    if (Test-Path -LiteralPath $Paths.Checkpoint) {
        $CheckpointData = Read-JsonFile $Paths.Checkpoint
        $Checkpoint = ConvertTo-UtcDate $CheckpointData.LastSuccessfulEndUtc
        Write-AuditLog "Last successful checkpoint: $($Checkpoint.ToString($IsoFormat, $Invariant)) (run $($CheckpointData.LastRunId))"
    }
    else {
        Write-AuditLog 'No checkpoint found. This is the first run for this data folder.'
    }

    # ---- Collection window (half-open: ge start, lt end)
    $NowUtc = (Get-Date).ToUniversalTime()
    $WindowEnd = New-Object DateTime($NowUtc.Year, $NowUtc.Month, $NowUtc.Day, $NowUtc.Hour, $NowUtc.Minute, 0, [System.DateTimeKind]::Utc)

    if ($Mode -eq 'Test') {
        $LookbackCap = $WindowEnd.AddHours(-$TestHours)
    }
    else {
        $LookbackCap = $WindowEnd.AddDays(-$MaxLookbackDays)
    }

    if ($null -ne $Checkpoint) {
        $WindowStart = $Checkpoint.AddHours(-$OverlapHours)
    }
    elseif ($Mode -eq 'Test') {
        $WindowStart = $LookbackCap
    }
    else {
        $WindowStart = $WindowEnd.AddDays(-$FirstRunDays)
    }

    if ($WindowStart -lt $LookbackCap) {
        if ($Mode -eq 'Production' -and $null -ne $Checkpoint -and $Checkpoint -lt $LookbackCap) {
            Write-AuditLog "POSSIBLE DATA GAP: checkpoint is older than $MaxLookbackDays days. Records between $($Checkpoint.ToString($IsoFormat, $Invariant)) and $($LookbackCap.ToString($IsoFormat, $Invariant)) may have expired from Entra retention." 'WARNING'
        }
        $WindowStart = $LookbackCap
    }

    if ($WindowStart -ge $WindowEnd) {
        throw "Invalid collection window: $WindowStart -> $WindowEnd"
    }

    Write-AuditLog "WINDOW START (UTC): $($WindowStart.ToString($IsoFormat, $Invariant))"
    Write-AuditLog "WINDOW END   (UTC): $($WindowEnd.ToString($IsoFormat, $Invariant))"
    Write-AuditLog ("WINDOW LENGTH: {0:N1} hours (overlap {1} h)" -f ($WindowEnd - $WindowStart).TotalHours, $OverlapHours)

    # ---- Graph authentication
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $SecretFilePath)) {
        throw 'Encrypted client secret file not found.'
    }

    $ClientSecret = Import-Clixml -LiteralPath $SecretFilePath
    if ($ClientSecret -isnot [System.Security.SecureString] -or $ClientSecret.Length -eq 0) {
        throw 'Encrypted client secret is invalid or empty.'
    }

    $Credential = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)

    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential -NoWelcome -ErrorAction Stop
    $GraphConnected = $true

    $GraphContext = Get-MgContext
    if ($GraphContext.AuthType -ne 'AppOnly' -or $GraphContext.ClientId -ne $ClientId) {
        throw 'Microsoft Graph AppOnly authentication was not established for the configured app.'
    }

    Write-AuditLog 'Microsoft Graph AppOnly authentication successful.'

    # ---- Process the window in chunks. Each chunk: collect BOTH streams into
    #      staging, commit them to the archive, then advance the checkpoint.
    #      An interruption (laptop closed) loses only the chunk in progress.
    New-Item -ItemType Directory -Path $RunStaging -Force | Out-Null

    $Committed = @{
        Interactive    = [PSCustomObject]@{ Staged = 0; New = 0; Skipped = 0 }
        NonInteractive = [PSCustomObject]@{ Staged = 0; New = 0; Skipped = 0 }
    }

    $ChunkCount = [int][Math]::Ceiling(($WindowEnd - $WindowStart).TotalHours / $ChunkHours)
    $ChunkNo = 0
    $ChunkStart = $WindowStart

    while ($ChunkStart -lt $WindowEnd) {
        $ChunkNo++
        $ChunkEnd = $ChunkStart.AddHours($ChunkHours)
        if ($ChunkEnd -gt $WindowEnd) { $ChunkEnd = $WindowEnd }

        Assert-NotStopRequested
        Write-AuditLog "---- Chunk $ChunkNo of $ChunkCount :$($ChunkStart.ToString($IsoFormat, $Invariant)) -> $($ChunkEnd.ToString($IsoFormat, $Invariant))"

        $script:ChunkStaging = Join-Path $RunStaging ('chunk_{0:D3}' -f $ChunkNo)
        New-Item -ItemType Directory -Path $script:ChunkStaging -Force | Out-Null

        # Phase 1: collect both streams into staging (nothing committed yet)
        foreach ($Stream in $Streams) {
            [void](Save-StreamToStaging -Stream $Stream -WindowStart $ChunkStart -WindowEnd $ChunkEnd)

            if ($SimulateFailure -eq 'AfterInteractive' -and $Stream.Name -eq 'Interactive') {
                throw 'SIMULATED FAILURE after interactive collection (test mode).'
            }
        }

        # Phase 2: commit to archive (idempotent - dedupe by Id) + user activity
        $ActivityState = Read-ActivityState $Paths
        foreach ($Stream in $Streams) {
            $C = Publish-StreamToArchive -StreamName $Stream.Name -ActivityState $ActivityState
            $Committed[$Stream.Name].Staged += $C.Staged
            $Committed[$Stream.Name].New += $C.New
            $Committed[$Stream.Name].Skipped += $C.Skipped
        }
        Save-ActivityState -Paths $Paths -State $ActivityState
        Write-AuditLog "User activity state saved ($($ActivityState.Users.Count) users)."

        if ($SimulateFailure -eq 'BeforeCheckpoint') {
            throw 'SIMULATED FAILURE after archive commit, before checkpoint (test mode).'
        }

        # Phase 3: advance checkpoint to the end of this chunk (LAST)
        $NewCheckpoint = [ordered]@{
            LastSuccessfulEndUtc  = $ChunkEnd.ToString($IsoFormat, $Invariant)
            LastWindowStartUtc    = $ChunkStart.ToString($IsoFormat, $Invariant)
            LastRunId             = $RunId
            LastRunCompletedLocal = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $Invariant)
            Mode                  = $Mode
            InteractiveNew        = $Committed['Interactive'].New
            NonInteractiveNew     = $Committed['NonInteractive'].New
        }

        # Never move the checkpoint backwards (overlap chunks re-read old hours)
        if ($null -ne $Checkpoint -and $ChunkEnd -le $Checkpoint) {
            Write-AuditLog "Chunk $ChunkNo of $ChunkCount saved (re-read of already-covered hours; checkpoint stays at $($Checkpoint.ToString($IsoFormat, $Invariant)))."
            Remove-Item -LiteralPath $script:ChunkStaging -Recurse -Force
            $ChunkStart = $ChunkEnd
            continue
        }

        Write-JsonFile -Object $NewCheckpoint -Path $Paths.Checkpoint -Depth 2
        $Checkpoint = $ChunkEnd

        Add-Content -LiteralPath $Paths.CheckpointLog -Encoding UTF8 -Value (
            '{0} | run {1} | {2} | chunk {3}/{4} {5} -> {6} | interactive +{7} | noninteractive +{8} (run totals)' -f `
                (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $RunId, $Trigger, $ChunkNo, $ChunkCount, `
                $NewCheckpoint.LastWindowStartUtc, $NewCheckpoint.LastSuccessfulEndUtc, `
                $NewCheckpoint.InteractiveNew, $NewCheckpoint.NonInteractiveNew
        )

        Write-AuditLog "Checkpoint advanced to $($NewCheckpoint.LastSuccessfulEndUtc) (chunk $ChunkNo of $ChunkCount saved)."

        Remove-Item -LiteralPath $script:ChunkStaging -Recurse -Force
        $ChunkStart = $ChunkEnd
    }

    # Graph no longer needed
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    $GraphConnected = $false
    Write-AuditLog 'Microsoft Graph disconnected.'

    foreach ($Name in $Committed.Keys) {
        if ($Committed[$Name].Staged -eq 0) {
            Write-AuditLog "$Name returned zero rows for the whole window. Verify this is expected." 'WARNING'
        }
    }

    # ---- Clean up this run's staging only after success
    Remove-Item -LiteralPath $RunStaging -Recurse -Force

    $Leftover = @(Get-ChildItem -LiteralPath $Paths.Staging -Directory -ErrorAction SilentlyContinue)
    if ($Leftover.Count -gt 0) {
        Write-AuditLog "$($Leftover.Count) staging folder(s) from earlier failed runs remain in $($Paths.Staging) (safe to review/delete)." 'WARNING'
    }

    Write-AuditLog '============================================'
    Write-AuditLog 'COLLECTION SUMMARY'
    foreach ($Stream in $Streams) {
        $C = $Committed[$Stream.Name]
        Write-AuditLog "$($Stream.Name): collected $($C.Staged), new $($C.New), duplicates skipped $($C.Skipped)"
    }
    Write-AuditLog "Graph pages: $script:PageCount, retries: $script:RetryCount, slice splits: $script:SplitCount"
    Write-AuditLog '============================================'

    $Succeeded = $true
}
catch {
    $FailureMessage = $_.Exception.Message
    if ($FailureMessage -eq $StopMessage) {
        $WasStopped = $true
        Write-AuditLog 'COLLECTION STOPPED BY USER.' 'WARNING'
        Write-AuditLog 'Days already saved are kept. The next run continues from the last checkpoint.' 'WARNING'
    }
    else {
        Write-AuditLog 'COLLECTION FAILED.' 'ERROR'
        Write-AuditLog "Error details: $FailureMessage" 'ERROR'
        Write-AuditLog 'Checkpoint NOT advanced. The next run will re-collect from the last successful checkpoint.' 'ERROR'
    }
    if (Test-Path -LiteralPath $RunStaging) {
        Write-AuditLog "Staged data kept for investigation: $RunStaging" 'WARNING'
    }
}
finally {
    if ($GraphConnected) {
        try {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
            Write-AuditLog 'Microsoft Graph disconnected.'
        }
        catch {
            Write-AuditLog "Graph disconnection warning: $($_.Exception.Message)" 'WARNING'
        }
    }

    # ---- Report rebuild (never changes the collection result)
    $ReportResult = $null
    if ($Succeeded) {
        try {
            $ReportResult = Invoke-EntraReportBuild -Paths $Paths
            Write-AuditLog "Report rebuilt: $($ReportResult.Users) users, $($ReportResult.RemoveCount) to remove, $($ReportResult.ReviewCount) to review. Licence file: $(if ($ReportResult.LicenseFile) { $ReportResult.LicenseFile } else { 'none uploaded yet' })"
        }
        catch {
            Write-AuditLog "Report rebuild failed (collection itself succeeded): $($_.Exception.Message)" 'WARNING'
        }
    }

    # ---- Status for dashboard + notification (only when this run held the lock)
    if ($null -ne $LockStream) {
        $Duration = [int]((Get-Date) - $StartedAt).TotalMinutes
        if ($Succeeded) {
            $Title = "Entra sign-in collection succeeded ($Mode)"
            $Message = "Interactive +$($Committed['Interactive'].New), non-interactive +$($Committed['NonInteractive'].New) new sign-ins in $Duration min."
            if ($ReportResult) { $Message += " Licences: $($ReportResult.RemoveCount) to remove, $($ReportResult.ReviewCount) to review." }
            $Level = 'success'
            $RunStatus = 'Success'
        }
        elseif ($WasStopped) {
            $Title = "Entra sign-in collection stopped ($Mode)"
            $Message = 'Stopped from the dashboard. Days already saved are kept; the next run continues from there.'
            $Level = 'warning'
            $RunStatus = 'Stopped'
        }
        else {
            $Title = "Entra sign-in collection FAILED ($Mode)"
            $Message = "$FailureMessage Checkpoint not advanced; next run will retry."
            $Level = 'error'
            $RunStatus = 'Failed'
        }

        if (Test-Path -LiteralPath $Paths.StopRequest) { Remove-Item -LiteralPath $Paths.StopRequest -Force -ErrorAction SilentlyContinue }

        try {
            Write-JsonFile -Path $Paths.LastRun -Object ([ordered]@{
                    Status = $RunStatus; RunId = $RunId; Trigger = $Trigger; Mode = $Mode
                    StartedLocal = $StartedAt.ToString('yyyy-MM-dd HH:mm:ss', $Invariant)
                    FinishedLocal = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $Invariant)
                    Message = $Message; LogFile = $LogFile
                })
            Add-AuditNotification -Paths $Paths -Level $Level -Title $Title -Message $Message
        }
        catch {
            Write-AuditLog "Could not write run status: $($_.Exception.Message)" 'WARNING'
        }

        if (-not $NoNotification) {
            if (Send-DesktopNotification -Title $Title -Message $Message) {
                Write-AuditLog 'Desktop notification shown.'
            }
            else {
                Write-AuditLog 'Desktop notification could not be shown.' 'WARNING'
            }
        }

        Set-KeepAwake $false
        $LockStream.Dispose()
    }

    if ($Succeeded) {
        Write-AuditLog 'FINAL STATUS: SUCCESS'
    }
    elseif ($WasStopped) {
        Write-AuditLog 'FINAL STATUS: STOPPED BY USER' 'WARNING'
    }
    else {
        Write-AuditLog 'FINAL STATUS: FAILED' 'ERROR'
    }

    Write-AuditLog "Execution log: $LogFile"
}

if ($Succeeded) { exit 0 } else { exit 1 }

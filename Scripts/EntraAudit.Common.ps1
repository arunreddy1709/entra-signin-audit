<#
============================================================
 ENTRA AUDIT - SHARED FUNCTIONS (Windows PowerShell 5.1)
 Dot-source from the other scripts:
     . (Join-Path $PSScriptRoot 'EntraAudit.Common.ps1')

 Contains: folder layout, date helpers, safe file publish,
 per-user activity state, licence review report builder,
 simple .xlsx writer, desktop notifications, run status.
============================================================
#>

$script:EntraIsoFormat = "yyyy-MM-dd'T'HH:mm:ss'Z'"
$script:EntraInvariant = [System.Globalization.CultureInfo]::InvariantCulture

# Licence review thresholds (days since last successful sign-in)
$script:EntraActiveDays = 30
$script:EntraReviewDays = 60
$script:EntraRemoveDays = 90

# ------------------------------------------------------------
# Folder layout
# ------------------------------------------------------------

function Get-AuditPaths {
    param(
        [ValidateSet('Test', 'Production')]
        [string]$Mode = 'Production'
    )

    $ScriptsFolder = $PSScriptRoot
    $ProjectRoot = Split-Path -Parent $ScriptsFolder

    if ($Mode -eq 'Test') {
        $DataRoot = Join-Path $ProjectRoot '06_Scheduler_Test\Weekly_Test'
    }
    else {
        $DataRoot = Join-Path $ProjectRoot '07_Weekly_Collector'
    }

    $State = Join-Path $DataRoot 'State'

    # Read-only copies for colleagues/managers (share this folder in OneDrive as "Can view")
    if ($Mode -eq 'Test') { $Shared = Join-Path $DataRoot 'Shared_Reports' }
    else { $Shared = Join-Path $ProjectRoot 'Shared_Reports' }

    [PSCustomObject]@{
        Mode           = $Mode
        ScriptsFolder  = $ScriptsFolder
        ProjectRoot    = $ProjectRoot
        DataRoot       = $DataRoot
        Archive        = Join-Path $DataRoot 'Archive'
        Staging        = Join-Path $DataRoot 'Staging'
        State          = $State
        Logs           = Join-Path $DataRoot 'Logs'
        Reports        = Join-Path $DataRoot 'Reports'
        Licenses       = Join-Path $DataRoot 'Licenses'
        Shared         = $Shared
        Checkpoint     = Join-Path $State 'checkpoint.json'
        CheckpointLog  = Join-Path $State 'checkpoint_history.log'
        Lock           = Join-Path $State 'collector.lock'
        StopRequest    = Join-Path $State 'stop.request'
        TaskName       = $(if ($Mode -eq 'Test') { 'Entra ID Daily Sign-In Collector - TEST' } else { 'Entra ID Daily Sign-In Collector' })
        ActivityState  = Join-Path $State 'user_activity.json'
        LastRun        = Join-Path $State 'last_run.json'
        Notifications  = Join-Path $State 'notifications.json'
        Collector      = Join-Path $ScriptsFolder 'Export-EntraSignIns_Weekly.ps1'
        Dashboard      = Join-Path $ProjectRoot 'Dashboard'
    }
}

function Initialize-AuditFolders {
    param($Paths)

    foreach ($Folder in @($Paths.DataRoot, $Paths.Archive, $Paths.Staging, $Paths.State, $Paths.Logs, $Paths.Reports, $Paths.Licenses)) {
        if (-not (Test-Path -LiteralPath $Folder)) {
            New-Item -ItemType Directory -Path $Folder -Force | Out-Null
        }
    }
}

# ------------------------------------------------------------
# Dates
# ------------------------------------------------------------

function ConvertTo-UtcDate {
    param($Value)

    if ($null -eq $Value -or [string]$Value -eq '') {
        throw 'Empty date value.'
    }

    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [System.DateTimeKind]::Local) { return $Value.ToUniversalTime() }
        if ($Value.Kind -eq [System.DateTimeKind]::Unspecified) { return [datetime]::SpecifyKind($Value, [System.DateTimeKind]::Utc) }
        return $Value
    }

    $Styles = [System.Globalization.DateTimeStyles]'AssumeUniversal, AdjustToUniversal'
    return [datetime]::Parse([string]$Value, $script:EntraInvariant, $Styles)
}

function ConvertTo-UtcIso {
    param($Value)
    return (ConvertTo-UtcDate $Value).ToString($script:EntraIsoFormat, $script:EntraInvariant)
}

function Format-LocalDate {
    param([string]$IsoUtc)
    if ([string]::IsNullOrEmpty($IsoUtc)) { return '' }
    return (ConvertTo-UtcDate $IsoUtc).ToLocalTime().ToString('yyyy-MM-dd HH:mm', $script:EntraInvariant)
}

# ------------------------------------------------------------
# Files
# ------------------------------------------------------------

function Publish-File {
    # Replace Destination with Source, keeping Destination.bak
    param([string]$Source, [string]$Destination)

    if (Test-Path -LiteralPath $Destination) {
        $Backup = "$Destination.bak"
        try {
            [System.IO.File]::Replace($Source, $Destination, $Backup)
        }
        catch {
            Copy-Item -LiteralPath $Destination -Destination $Backup -Force
            Move-Item -LiteralPath $Source -Destination $Destination -Force
        }
    }
    else {
        Move-Item -LiteralPath $Source -Destination $Destination -Force
    }
}

function Write-JsonFile {
    param($Object, [string]$Path, [int]$Depth = 6)

    $Temp = "$Path.tmp"
    $Json = ConvertTo-Json -InputObject $Object -Depth $Depth -Compress
    [System.IO.File]::WriteAllText($Temp, $Json, (New-Object System.Text.UTF8Encoding($false)))
    Publish-File -Source $Temp -Destination $Path
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $Text = [System.IO.File]::ReadAllText($Path)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    return ($Text | ConvertFrom-Json)
}

# ------------------------------------------------------------
# Per-user activity state (last successful sign-ins)
# ------------------------------------------------------------

function Read-ActivityState {
    param($Paths)

    $State = @{
        HistoryStartUtc = ''
        Users           = @{}
    }

    $Data = Read-JsonFile $Paths.ActivityState
    if ($null -ne $Data) {
        $State.HistoryStartUtc = [string]$Data.HistoryStartUtc
        foreach ($U in @($Data.Users)) {
            if ($null -eq $U -or [string]::IsNullOrEmpty($U.UserId)) { continue }
            $State.Users[[string]$U.UserId] = @{
                UserId             = [string]$U.UserId
                Upn                = [string]$U.Upn
                Name               = [string]$U.Name
                LastInteractive    = [string]$U.LastInteractive
                LastNonInteractive = [string]$U.LastNonInteractive
            }
        }
    }

    return $State
}

function Update-ActivityState {
    # Keeps the latest successful sign-in per user and kind. Idempotent.
    param(
        $State,
        [string]$UserId,
        [string]$Upn,
        [string]$Name,
        [ValidateSet('Interactive', 'NonInteractive')][string]$Kind,
        [string]$IsoUtc
    )

    if ([string]::IsNullOrEmpty($UserId) -or [string]::IsNullOrEmpty($IsoUtc)) { return }

    if (-not $State.Users.ContainsKey($UserId)) {
        $State.Users[$UserId] = @{
            UserId = $UserId; Upn = $Upn; Name = $Name
            LastInteractive = ''; LastNonInteractive = ''
        }
    }

    $Entry = $State.Users[$UserId]
    $Field = "Last$Kind"

    # ISO-8601 UTC strings of identical format compare correctly as text
    if ([string]::CompareOrdinal($IsoUtc, [string]$Entry[$Field]) -gt 0) {
        $Entry[$Field] = $IsoUtc
        if ($Upn) { $Entry.Upn = $Upn }
        if ($Name) { $Entry.Name = $Name }
    }

    if ([string]::IsNullOrEmpty($State.HistoryStartUtc) -or [string]::CompareOrdinal($IsoUtc, $State.HistoryStartUtc) -lt 0) {
        $State.HistoryStartUtc = $IsoUtc
    }
}

function Save-ActivityState {
    param($Paths, $State)

    $Users = @($State.Users.Values | Sort-Object { $_.Upn } | ForEach-Object {
            [ordered]@{
                UserId             = $_.UserId
                Upn                = $_.Upn
                Name               = $_.Name
                LastInteractive    = $_.LastInteractive
                LastNonInteractive = $_.LastNonInteractive
            }
        })

    $Object = [ordered]@{
        UpdatedUtc      = (Get-Date).ToUniversalTime().ToString($script:EntraIsoFormat, $script:EntraInvariant)
        HistoryStartUtc = $State.HistoryStartUtc
        Users           = $Users
    }

    Write-JsonFile -Object $Object -Path $Paths.ActivityState -Depth 4
}

# ------------------------------------------------------------
# Run status and notifications
# ------------------------------------------------------------

function Add-AuditNotification {
    param($Paths, [string]$Level, [string]$Title, [string]$Message)

    $List = New-Object System.Collections.ArrayList
    $Existing = Read-JsonFile $Paths.Notifications
    if ($null -ne $Existing) { foreach ($N in @($Existing)) { [void]$List.Add($N) } }

    [void]$List.Insert(0, [ordered]@{
            TimeLocal = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:EntraInvariant)
            Level     = $Level
            Title     = $Title
            Message   = $Message
        })

    while ($List.Count -gt 50) { $List.RemoveAt($List.Count - 1) }

    Write-JsonFile -Object @($List) -Path $Paths.Notifications -Depth 3
}

function Set-KeepAwake {
    # Stops Windows going to sleep while a collection runs (lid-close
    # behaviour still follows the power settings). Never throws.
    param([bool]$On)
    try {
        if (-not ('EntraAudit.Power' -as [type])) {
            Add-Type -Namespace EntraAudit -Name Power -MemberDefinition '[DllImport("kernel32.dll")] public static extern uint SetThreadExecutionState(uint esFlags);'
        }
        # ES_CONTINUOUS (0x80000000) | ES_SYSTEM_REQUIRED (0x00000001)
        if ($On) { [void][EntraAudit.Power]::SetThreadExecutionState([uint32]2147483649) }
        else { [void][EntraAudit.Power]::SetThreadExecutionState([uint32]2147483648) }
    }
    catch { }
}

function Send-DesktopNotification {
    # Windows toast notification. Never throws.
    param([string]$Title, [string]$Message)

    try {
        [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]

        $T = [System.Security.SecurityElement]::Escape($Title)
        $M = [System.Security.SecurityElement]::Escape($Message)

        $Xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $Xml.LoadXml("<toast><visual><binding template='ToastGeneric'><text>$T</text><text>$M</text></binding></visual></toast>")

        $Toast = New-Object Windows.UI.Notifications.ToastNotification($Xml)
        $AppId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId).Show($Toast)
        return $true
    }
    catch {
        return $false
    }
}

# ------------------------------------------------------------
# Simple .xlsx writer (no Excel needed)
# ------------------------------------------------------------

function Get-XlsxColumnName {
    param([int]$Index)  # 1-based
    $Name = ''
    while ($Index -gt 0) {
        $Mod = ($Index - 1) % 26
        $Name = [char](65 + $Mod) + $Name
        $Index = [int][Math]::Floor(($Index - $Mod) / 26)
    }
    return $Name
}

function ConvertTo-XmlText {
    param($Value)
    $Text = [string]$Value
    $Text = [regex]::Replace($Text, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    if ($Text.Length -gt 32000) { $Text = $Text.Substring(0, 32000) }
    return [System.Security.SecurityElement]::Escape($Text)
}

function Export-SimpleXlsx {
    <#
      Sheets: array of @{ Name = 'Sheet'; Columns = @('A','B'); Rows = @(objects); Widths = @{ 'A' = 30 } }
      Numbers stay numbers; everything else is text. Header row is bold,
      frozen and has an AutoFilter.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][array]$Sheets
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $Utf8 = New-Object System.Text.UTF8Encoding($false)
    $Temp = "$Path.tmp"
    if (Test-Path -LiteralPath $Temp) { Remove-Item -LiteralPath $Temp -Force }

    $FileStream = [System.IO.File]::Open($Temp, [System.IO.FileMode]::CreateNew)
    $Zip = New-Object System.IO.Compression.ZipArchive($FileStream, [System.IO.Compression.ZipArchiveMode]::Create)

    function Add-ZipText([string]$EntryName, [string]$Text) {
        $Entry = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
        $Writer = New-Object System.IO.StreamWriter($Entry.Open(), $Utf8)
        $Writer.Write($Text)
        $Writer.Dispose()
    }

    try {
        $Main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        $RelNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

        $ContentTypes = New-Object System.Text.StringBuilder
        [void]$ContentTypes.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')

        $WorkbookSheets = New-Object System.Text.StringBuilder
        $DefinedNames = New-Object System.Text.StringBuilder
        $WorkbookRels = New-Object System.Text.StringBuilder
        [void]$WorkbookRels.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">")

        for ($S = 0; $S -lt $Sheets.Count; $S++) {
            $Sheet = $Sheets[$S]
            $SheetNo = $S + 1
            $Columns = @($Sheet.Columns)
            $Rows = @($Sheet.Rows)
            $LastCol = Get-XlsxColumnName $Columns.Count
            $LastRow = [Math]::Max(1, $Rows.Count + 1)
            $SafeName = ($Sheet.Name -replace '[\[\]\*\?/\\:]', '-')
            if ($SafeName.Length -gt 31) { $SafeName = $SafeName.Substring(0, 31) }

            $Sb = New-Object System.Text.StringBuilder
            [void]$Sb.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet xmlns=`"$Main`" xmlns:r=`"$RelNs`">")
            [void]$Sb.Append("<dimension ref=`"A1:$LastCol$LastRow`"/>")
            [void]$Sb.Append('<sheetViews><sheetView workbookViewId="0"' + $(if ($S -eq 0) { ' tabSelected="1"' } else { '' }) + '><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')

            [void]$Sb.Append('<cols>')
            for ($C = 0; $C -lt $Columns.Count; $C++) {
                $Width = 18
                if ($Sheet.Widths -and $Sheet.Widths.ContainsKey($Columns[$C])) { $Width = $Sheet.Widths[$Columns[$C]] }
                [void]$Sb.Append("<col min=`"$($C + 1)`" max=`"$($C + 1)`" width=`"$Width`" customWidth=`"1`"/>")
            }
            [void]$Sb.Append('</cols><sheetData>')

            # Header
            [void]$Sb.Append('<row r="1">')
            for ($C = 0; $C -lt $Columns.Count; $C++) {
                $Ref = (Get-XlsxColumnName ($C + 1)) + '1'
                [void]$Sb.Append("<c r=`"$Ref`" t=`"inlineStr`" s=`"1`"><is><t>$(ConvertTo-XmlText $Columns[$C])</t></is></c>")
            }
            [void]$Sb.Append('</row>')

            # Data
            for ($R = 0; $R -lt $Rows.Count; $R++) {
                $RowNo = $R + 2
                $Item = $Rows[$R]
                [void]$Sb.Append("<row r=`"$RowNo`">")
                for ($C = 0; $C -lt $Columns.Count; $C++) {
                    $Col = $Columns[$C]
                    if ($Item -is [System.Collections.IDictionary]) { $Value = $Item[$Col] } else { $Value = $Item.$Col }
                    if ($null -eq $Value -or [string]$Value -eq '') { continue }
                    $Ref = (Get-XlsxColumnName ($C + 1)) + $RowNo
                    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
                        [void]$Sb.Append("<c r=`"$Ref`"><v>$([string]::Format($script:EntraInvariant, '{0}', $Value))</v></c>")
                    }
                    else {
                        [void]$Sb.Append("<c r=`"$Ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText $Value)</t></is></c>")
                    }
                }
                [void]$Sb.Append('</row>')
            }

            [void]$Sb.Append("</sheetData><autoFilter ref=`"A1:$LastCol$LastRow`"/><pageMargins left=`"0.7`" right=`"0.7`" top=`"0.75`" bottom=`"0.75`" header=`"0.3`" footer=`"0.3`"/></worksheet>")

            Add-ZipText "xl/worksheets/sheet$SheetNo.xml" $Sb.ToString()

            [void]$ContentTypes.Append("<Override PartName=`"/xl/worksheets/sheet$SheetNo.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>")
            [void]$WorkbookSheets.Append("<sheet name=`"$(ConvertTo-XmlText $SafeName)`" sheetId=`"$SheetNo`" r:id=`"rId$SheetNo`"/>")
            [void]$DefinedNames.Append("<definedName name=`"_xlnm._FilterDatabase`" localSheetId=`"$S`" hidden=`"1`">'$(ConvertTo-XmlText ($SafeName -replace "'", "''"))'!`$A`$1:`$$LastCol`$$LastRow</definedName>")
            [void]$WorkbookRels.Append("<Relationship Id=`"rId$SheetNo`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$SheetNo.xml`"/>")
        }

        $StylesId = $Sheets.Count + 1
        [void]$WorkbookRels.Append("<Relationship Id=`"rId$StylesId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/></Relationships>")
        [void]$ContentTypes.Append('</Types>')

        Add-ZipText '[Content_Types].xml' $ContentTypes.ToString()
        Add-ZipText '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        Add-ZipText 'xl/workbook.xml' "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><workbook xmlns=`"$Main`" xmlns:r=`"$RelNs`"><bookViews><workbookView/></bookViews><sheets>$($WorkbookSheets.ToString())</sheets><definedNames>$($DefinedNames.ToString())</definedNames></workbook>"
        Add-ZipText 'xl/_rels/workbook.xml.rels' $WorkbookRels.ToString()
        Add-ZipText 'xl/styles.xml' "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><styleSheet xmlns=`"$Main`"><fonts count=`"2`"><font><sz val=`"11`"/><name val=`"Calibri`"/><family val=`"2`"/></font><font><b/><sz val=`"11`"/><color rgb=`"FFFFFFFF`"/><name val=`"Calibri`"/><family val=`"2`"/></font></fonts><fills count=`"3`"><fill><patternFill patternType=`"none`"/></fill><fill><patternFill patternType=`"gray125`"/></fill><fill><patternFill patternType=`"solid`"><fgColor rgb=`"FF1C5CAB`"/><bgColor indexed=`"64`"/></patternFill></fill></fills><borders count=`"1`"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=`"1`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`"/></cellStyleXfs><cellXfs count=`"2`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`" xfId=`"0`"/><xf numFmtId=`"0`" fontId=`"1`" fillId=`"2`" borderId=`"0`" xfId=`"0`" applyFont=`"1`" applyFill=`"1`"/></cellXfs><cellStyles count=`"1`"><cellStyle name=`"Normal`" xfId=`"0`" builtinId=`"0`"/></cellStyles></styleSheet>"
    }
    finally {
        $Zip.Dispose()
        $FileStream.Dispose()
    }

    Publish-File -Source $Temp -Destination $Path
    if (Test-Path -LiteralPath "$Path.bak") { Remove-Item -LiteralPath "$Path.bak" -Force -ErrorAction SilentlyContinue }
}

# ------------------------------------------------------------
# Licence review report
# ------------------------------------------------------------

$script:ReportColumns = @(
    'id', 'displayName', 'userPrincipalName', 'userType',
    'LastInteractiveSignIn', 'DaysSinceLastSignIn', 'ActivityStatus',
    'LastNonInteractiveSignIn', 'AccountStatus', 'Licenses', 'Block credential',
    'LastRecordedActivity', 'DaysSinceLastActivity', 'CombinedActivityStatus',
    'LicenseStatus', 'InactivityCategory', 'LicenseReviewStatus',
    'LicenseCount', 'PaidLicenseCount', 'RecommendedAction'
)

function Test-FreeLicense {
    param([string]$Name)
    return ($Name -match '\bFree\b|Trial|vTrial|for Developer|Exploratory')
}

function Get-LatestLicenseFile {
    param($Paths)
    if (-not (Test-Path -LiteralPath $Paths.Licenses)) { return $null }
    return Get-ChildItem -LiteralPath $Paths.Licenses -Filter '*.csv' -File |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Test-LicenseCsvHeader {
    # Returns $null if OK, otherwise an error message
    param([string]$HeaderLine)
    $Required = @('User principal name', 'Licenses', 'Object Id', 'Block credential')
    $Missing = @($Required | Where-Object { $HeaderLine -notmatch [regex]::Escape($_) })
    if ($Missing.Count -gt 0) {
        return "This does not look like a Microsoft 365 admin center 'Active users' export. Missing column(s): $($Missing -join ', ')."
    }
    return $null
}

function Get-DaysSince {
    param([string]$IsoUtc, [datetime]$NowUtc)
    if ([string]::IsNullOrEmpty($IsoUtc)) { return $null }
    return [int][Math]::Floor(($NowUtc - (ConvertTo-UtcDate $IsoUtc)).TotalDays)
}

function New-UserReportRow {
    param(
        [string]$Id, [string]$Name, [string]$Upn, [string]$UserType,
        [string]$Licenses, [string]$BlockCredential, [bool]$InLicenseFile,
        $Activity, [datetime]$NowUtc
    )

    $LastI = ''; $LastN = ''
    if ($null -ne $Activity) { $LastI = [string]$Activity.LastInteractive; $LastN = [string]$Activity.LastNonInteractive }

    $LastAny = $LastI
    if ([string]::CompareOrdinal($LastN, $LastAny) -gt 0) { $LastAny = $LastN }

    $DaysI = Get-DaysSince $LastI $NowUtc
    $DaysAny = Get-DaysSince $LastAny $NowUtc

    if ($null -eq $DaysI) { $ActivityStatus = 'No successful interactive sign-in recorded' }
    elseif ($DaysI -le $script:EntraActiveDays) { $ActivityStatus = 'Recent interactive activity' }
    else { $ActivityStatus = "No interactive sign-in in last $($script:EntraActiveDays) days" }

    if ($null -eq $DaysAny) { $Combined = 'No successful sign-in recorded' }
    elseif ($DaysAny -le $script:EntraActiveDays) { $Combined = 'Recent recorded activity' }
    else { $Combined = "No recorded activity in last $($script:EntraActiveDays) days" }

    if ($null -eq $DaysAny) { $Inactivity = 'No recorded sign-in' }
    elseif ($DaysAny -le $script:EntraActiveDays) { $Inactivity = "Active within $($script:EntraActiveDays) days" }
    elseif ($DaysAny -le $script:EntraReviewDays) { $Inactivity = "Inactive $($script:EntraActiveDays + 1)-$($script:EntraReviewDays) days" }
    elseif ($DaysAny -le $script:EntraRemoveDays) { $Inactivity = "Inactive $($script:EntraReviewDays + 1)-$($script:EntraRemoveDays) days" }
    else { $Inactivity = "Inactive over $($script:EntraRemoveDays) days" }

    if (-not $InLicenseFile) { $AccountStatus = 'Not in licence export' }
    elseif ($BlockCredential -eq 'True') { $AccountStatus = 'Disabled' }
    else { $AccountStatus = 'Enabled' }

    $LicenseList = @()
    if ($Licenses -and $Licenses -ne 'Unlicensed') {
        $LicenseList = @($Licenses -split '\+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    $PaidCount = @($LicenseList | Where-Object { -not (Test-FreeLicense $_) }).Count

    if (-not $InLicenseFile) { $LicenseStatus = 'Not in licence export' }
    elseif ($LicenseList.Count -gt 0) { $LicenseStatus = 'Licensed' }
    else { $LicenseStatus = 'Unlicensed' }

    if ($LicenseStatus -ne 'Licensed') {
        $Review = $LicenseStatus
        $Action = ''
    }
    elseif ($AccountStatus -eq 'Disabled') {
        $Review = 'Remove - Licensed but sign-in blocked'
        $Action = 'Remove licences (account is blocked)'
    }
    elseif ($null -eq $DaysAny) {
        $Review = 'Review - Licensed with no recorded sign-in'
        $Action = 'Confirm owner; remove licences if unused'
    }
    elseif ($DaysAny -gt $script:EntraRemoveDays) {
        $Review = "Remove - Licensed, inactive over $($script:EntraRemoveDays) days"
        $Action = 'Remove licences'
    }
    elseif ($DaysAny -gt $script:EntraActiveDays) {
        $Review = "Review - Licensed, inactive over $($script:EntraActiveDays) days"
        $Action = 'Check with manager before removing'
    }
    else {
        $Review = 'Licensed - Activity found'
        $Action = ''
    }

    if ($PaidCount -eq 0 -and $Action) {
        $Action = "$Action (free/trial licences only)"
    }

    [ordered]@{
        'id'                       = $Id
        'displayName'              = $Name
        'userPrincipalName'        = $Upn
        'userType'                 = $UserType
        'LastInteractiveSignIn'    = Format-LocalDate $LastI
        'DaysSinceLastSignIn'      = $DaysI
        'ActivityStatus'           = $ActivityStatus
        'LastNonInteractiveSignIn' = Format-LocalDate $LastN
        'AccountStatus'            = $AccountStatus
        'Licenses'                 = $(if ($LicenseList.Count -gt 0) { $LicenseList -join '+' } else { 'Unlicensed' })
        'Block credential'         = $BlockCredential
        'LastRecordedActivity'     = Format-LocalDate $LastAny
        'DaysSinceLastActivity'    = $DaysAny
        'CombinedActivityStatus'   = $Combined
        'LicenseStatus'            = $LicenseStatus
        'InactivityCategory'       = $Inactivity
        'LicenseReviewStatus'      = $Review
        'LicenseCount'             = $LicenseList.Count
        'PaidLicenseCount'         = $PaidCount
        'RecommendedAction'        = $Action
    }
}

function Get-CountMap {
    param($Rows, [string]$Field)
    $Map = [ordered]@{}
    foreach ($G in ($Rows | Group-Object { $_[$Field] } | Sort-Object Count -Descending)) {
        $Map[[string]$G.Name] = $G.Count
    }
    return $Map
}

function New-SnapshotHtml {
    # The dashboard page with the report embedded: no server, no actions.
    param($Paths, $Report)

    $IndexPath = Join-Path $Paths.Dashboard 'index.html'
    $Template = [System.IO.File]::ReadAllText($IndexPath)

    $Json = ConvertTo-Json -InputObject $Report -Depth 6 -Compress
    $Json = $Json.Replace('</', '<\/')   # cannot close the <script> tag early

    $Html = $Template.Replace('__AUTH_TOKEN__', '').Replace('__MODE__', 'Snapshot')
    $Html = $Html.Replace('<!--ENTRA_SNAPSHOT-->', "<script>window.ENTRA_SNAPSHOT = $Json;</script>")
    return $Html
}

function Publish-SharedReports {
    <#
      Copies read-only outputs into the shared folder:
        Shared_Reports\Entra_Dashboard_View.html        (latest, opens offline)
        Shared_Reports\Entra_User_Activity_Latest.xlsx  (latest)
        Shared_Reports\Monthly\yyyy-MM\...              (one per month, kept)
      Files are overwritten in place so existing OneDrive share links keep working.
    #>
    param($Paths, $Report, [string]$MonthXlsx)

    $Month = $Report.Month
    $MonthShared = Join-Path (Join-Path $Paths.Shared 'Monthly') $Month
    foreach ($F in @($Paths.Shared, $MonthShared)) {
        if (-not (Test-Path -LiteralPath $F)) { New-Item -ItemType Directory -Path $F -Force | Out-Null }
    }

    $Utf8 = New-Object System.Text.UTF8Encoding($false)
    $Html = New-SnapshotHtml -Paths $Paths -Report $Report

    $TempHtml = Join-Path $Paths.Reports 'snapshot.tmp.html'
    [System.IO.File]::WriteAllText($TempHtml, $Html, $Utf8)

    Copy-Item -LiteralPath $TempHtml -Destination (Join-Path $Paths.Shared 'Entra_Dashboard_View.html') -Force
    Copy-Item -LiteralPath $TempHtml -Destination (Join-Path $MonthShared "Entra_Dashboard_View_$Month.html") -Force
    Remove-Item -LiteralPath $TempHtml -Force

    Copy-Item -LiteralPath $MonthXlsx -Destination (Join-Path $Paths.Shared 'Entra_User_Activity_Latest.xlsx') -Force
    Copy-Item -LiteralPath $MonthXlsx -Destination (Join-Path $MonthShared "Entra_User_Activity_$Month.xlsx") -Force

    $Readme = Join-Path $Paths.Shared 'README.txt'
    if (-not (Test-Path -LiteralPath $Readme)) {
        $Text = @(
            'ENTRA ID USER ACTIVITY & LICENCE REVIEW - READ-ONLY REPORTS'
            ''
            'Entra_User_Activity_Latest.xlsx'
            '    Latest report. Opens in Excel in the browser. Sheets: Summary,'
            '    Removal Candidates, FinalUserActivity, Licence Summary.'
            ''
            'Entra_Dashboard_View.html'
            '    Interactive read-only dashboard (charts, filters, tables).'
            '    Download it, then double-click to open in Edge or Chrome.'
            '    It works offline; nothing is sent anywhere.'
            ''
            'Monthly\yyyy-MM\'
            '    The saved report for each month.'
            ''
            'Updated automatically after each sign-in collection.'
            'Contains staff names, sign-in dates and licence data - internal use only.'
        ) -join "`r`n"
        [System.IO.File]::WriteAllText($Readme, $Text, $Utf8)
    }
}

function Invoke-EntraReportBuild {
    <#
      Builds the user activity / licence review report from:
        - State\user_activity.json  (filled by the collector)
        - newest CSV in Licenses\    (admin center Active users export)
      Writes Reports\latest.json, Reports\latest.xlsx and the current
      month's snapshot Reports\yyyy-MM\ (json + xlsx).
      Returns a summary object.
    #>
    param($Paths)

    Initialize-AuditFolders $Paths

    $NowUtc = (Get-Date).ToUniversalTime()
    $Activity = Read-ActivityState $Paths
    $LicenseFile = Get-LatestLicenseFile $Paths

    $Rows = New-Object System.Collections.Generic.List[object]
    $SeenIds = New-Object 'System.Collections.Generic.HashSet[string]'

    if ($null -ne $LicenseFile) {
        foreach ($U in (Import-Csv -LiteralPath $LicenseFile.FullName)) {
            $Id = [string]$U.'Object Id'
            if ([string]::IsNullOrEmpty($Id)) { continue }
            [void]$SeenIds.Add($Id)

            $Upn = [string]$U.'User principal name'
            $UserType = [string]$U.MemberType
            if ([string]::IsNullOrEmpty($UserType)) {
                if ($Upn -match '#EXT#') { $UserType = 'Guest' } else { $UserType = 'Member' }
            }

            $Rows.Add((New-UserReportRow -Id $Id -Name $U.'Display name' -Upn $Upn -UserType $UserType `
                        -Licenses $U.Licenses -BlockCredential $U.'Block credential' -InLicenseFile $true `
                        -Activity $Activity.Users[$Id] -NowUtc $NowUtc))
        }
    }

    # Users seen signing in but missing from the licence export
    foreach ($A in $Activity.Users.Values) {
        if ($SeenIds.Contains($A.UserId)) { continue }
        if ($A.Upn -match '#EXT#') { $UserType = 'Guest' } else { $UserType = 'Unknown' }
        $Rows.Add((New-UserReportRow -Id $A.UserId -Name $A.Name -Upn $A.Upn -UserType $UserType `
                    -Licenses '' -BlockCredential '' -InLicenseFile $false -Activity $A -NowUtc $NowUtc))
    }

    $RowArray = @($Rows | Sort-Object { $_['displayName'] })

    # Licence distribution: assigned vs assigned-to-inactive users
    $LicenseStats = @{}
    foreach ($R in $RowArray) {
        if ($R['LicenseStatus'] -ne 'Licensed') { continue }
        $Inactive = ($R['LicenseReviewStatus'] -notlike 'Licensed - Activity found')
        foreach ($L in ($R['Licenses'] -split '\+')) {
            if (-not $L) { continue }
            if (-not $LicenseStats.ContainsKey($L)) { $LicenseStats[$L] = @{ Name = $L; Assigned = 0; ToReview = 0; Free = (Test-FreeLicense $L) } }
            $LicenseStats[$L].Assigned++
            if ($Inactive) { $LicenseStats[$L].ToReview++ }
        }
    }
    $LicenseList = @($LicenseStats.Values | Sort-Object @{ e = { $_.Assigned }; Descending = $true }, @{ e = { $_.Name } } | ForEach-Object {
            [ordered]@{ Name = $_.Name; Assigned = $_.Assigned; ToReview = $_.ToReview; Free = [bool]$_.Free }
        })

    $Candidates = @($RowArray | Where-Object { $_['LicenseReviewStatus'] -like 'Remove*' -or $_['LicenseReviewStatus'] -like 'Review*' })
    $RemoveCount = @($Candidates | Where-Object { $_['LicenseReviewStatus'] -like 'Remove*' }).Count

    $CoverageDays = 0
    if ($Activity.HistoryStartUtc) { $CoverageDays = [int][Math]::Floor(($NowUtc - (ConvertTo-UtcDate $Activity.HistoryStartUtc)).TotalDays) }

    $LastRun = Read-JsonFile $Paths.LastRun
    $Month = (Get-Date).ToString('yyyy-MM', $script:EntraInvariant)

    $Summary = [ordered]@{
        TotalUsers         = $RowArray.Count
        Licensed           = @($RowArray | Where-Object { $_['LicenseStatus'] -eq 'Licensed' }).Count
        Unlicensed         = @($RowArray | Where-Object { $_['LicenseStatus'] -eq 'Unlicensed' }).Count
        NotInLicenseExport = @($RowArray | Where-Object { $_['LicenseStatus'] -eq 'Not in licence export' }).Count
        Active             = @($RowArray | Where-Object { $_['InactivityCategory'] -like 'Active*' }).Count
        Disabled           = @($RowArray | Where-Object { $_['AccountStatus'] -eq 'Disabled' }).Count
        Candidates         = $Candidates.Count
        RemoveCount        = $RemoveCount
        ReviewCount        = $Candidates.Count - $RemoveCount
        ByInactivity       = Get-CountMap $RowArray 'InactivityCategory'
        ByReview           = Get-CountMap $RowArray 'LicenseReviewStatus'
        ByAccount          = Get-CountMap $RowArray 'AccountStatus'
        ByLicenseStatus    = Get-CountMap $RowArray 'LicenseStatus'
        ByUserType         = Get-CountMap $RowArray 'userType'
    }

    $Report = [ordered]@{
        Month              = $Month
        GeneratedLocal     = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:EntraInvariant)
        GeneratedUtc       = $NowUtc.ToString($script:EntraIsoFormat, $script:EntraInvariant)
        Mode               = $Paths.Mode
        LicenseFile        = $(if ($LicenseFile) { $LicenseFile.Name } else { '' })
        LicenseFileDate    = $(if ($LicenseFile) { $LicenseFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm', $script:EntraInvariant) } else { '' })
        HistoryStartLocal  = Format-LocalDate $Activity.HistoryStartUtc
        CoverageDays       = $CoverageDays
        Thresholds         = [ordered]@{ ActiveDays = $script:EntraActiveDays; ReviewDays = $script:EntraReviewDays; RemoveDays = $script:EntraRemoveDays }
        LastRun            = $LastRun
        Summary            = $Summary
        LicenseStats       = $LicenseList
        Columns            = $script:ReportColumns
        Users              = $RowArray
    }

    # ---- Excel workbook
    $SummaryRows = @(
        [ordered]@{ Item = 'Report month'; Value = $Month }
        [ordered]@{ Item = 'Generated'; Value = $Report.GeneratedLocal }
        [ordered]@{ Item = 'Licence export used'; Value = $Report.LicenseFile }
        [ordered]@{ Item = 'Licence export date'; Value = $Report.LicenseFileDate }
        [ordered]@{ Item = 'Sign-in history since'; Value = $Report.HistoryStartLocal }
        [ordered]@{ Item = 'Sign-in history (days)'; Value = $CoverageDays }
        [ordered]@{ Item = 'Total users'; Value = $Summary.TotalUsers }
        [ordered]@{ Item = 'Licensed users'; Value = $Summary.Licensed }
        [ordered]@{ Item = 'Unlicensed users'; Value = $Summary.Unlicensed }
        [ordered]@{ Item = "Active within $($script:EntraActiveDays) days"; Value = $Summary.Active }
        [ordered]@{ Item = 'Disabled accounts'; Value = $Summary.Disabled }
        [ordered]@{ Item = 'Licence removal - REMOVE'; Value = $Summary.RemoveCount }
        [ordered]@{ Item = 'Licence removal - REVIEW'; Value = $Summary.ReviewCount }
    )
    foreach ($K in $Summary.ByReview.Keys) { $SummaryRows += [ordered]@{ Item = "Review status: $K"; Value = $Summary.ByReview[$K] } }
    foreach ($K in $Summary.ByInactivity.Keys) { $SummaryRows += [ordered]@{ Item = "Inactivity: $K"; Value = $Summary.ByInactivity[$K] } }

    $Widths = @{ 'displayName' = 28; 'userPrincipalName' = 36; 'Licenses' = 60; 'ActivityStatus' = 34; 'CombinedActivityStatus' = 34; 'LicenseReviewStatus' = 42; 'InactivityCategory' = 24; 'RecommendedAction' = 40; 'id' = 38 }

    $CandidateColumns = @('displayName', 'userPrincipalName', 'AccountStatus', 'LicenseReviewStatus', 'RecommendedAction', 'LastRecordedActivity', 'DaysSinceLastActivity', 'LicenseCount', 'PaidLicenseCount', 'Licenses', 'id')

    $Sheets = @(
        @{ Name = 'Summary'; Columns = @('Item', 'Value'); Rows = $SummaryRows; Widths = @{ Item = 40; Value = 40 } }
        @{ Name = 'Removal Candidates'; Columns = $CandidateColumns; Rows = $Candidates; Widths = $Widths }
        @{ Name = 'FinalUserActivity'; Columns = $script:ReportColumns; Rows = $RowArray; Widths = $Widths }
        @{ Name = 'Licence Summary'; Columns = @('Name', 'Assigned', 'ToReview', 'Free'); Rows = @($LicenseList | ForEach-Object { [ordered]@{ Name = $_.Name; Assigned = $_.Assigned; ToReview = $_.ToReview; Free = $(if ($_.Free) { 'Yes' } else { 'No' }) } }); Widths = @{ Name = 50 } }
    )

    $MonthFolder = Join-Path $Paths.Reports $Month
    if (-not (Test-Path -LiteralPath $MonthFolder)) { New-Item -ItemType Directory -Path $MonthFolder -Force | Out-Null }

    $MonthXlsx = Join-Path $MonthFolder "Entra_User_Activity_$Month.xlsx"
    Export-SimpleXlsx -Path $MonthXlsx -Sheets $Sheets
    Copy-Item -LiteralPath $MonthXlsx -Destination (Join-Path $Paths.Reports 'latest.xlsx') -Force

    Write-JsonFile -Object $Report -Path (Join-Path $MonthFolder 'report.json') -Depth 6
    Write-JsonFile -Object $Report -Path (Join-Path $Paths.Reports 'latest.json') -Depth 6

    Publish-SharedReports -Paths $Paths -Report $Report -MonthXlsx $MonthXlsx

    [PSCustomObject]@{
        Month       = $Month
        Users       = $Summary.TotalUsers
        Candidates  = $Summary.Candidates
        RemoveCount = $Summary.RemoveCount
        ReviewCount = $Summary.ReviewCount
        LicenseFile = $Report.LicenseFile
        Xlsx        = $MonthXlsx
    }
}

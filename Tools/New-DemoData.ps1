<#
 Generates FICTIONAL demo data in the Test data folder so the dashboard,
 Excel report and shared HTML view can be explored (and screenshotted)
 without a tenant. All names are invented; all accounts use example.com.

   .\Tools\New-DemoData.ps1
   .\Scripts\Start-EntraDashboard.ps1 -Mode Test

 Never run this against Production data.
#>
[CmdletBinding()]
param(
    [int]$UserCount = 80,
    [int]$Seed = 42
)

$ErrorActionPreference = 'Stop'
$Scripts = Join-Path (Split-Path -Parent $PSScriptRoot) 'Scripts'
. (Join-Path $Scripts 'EntraAudit.Common.ps1')

$Paths = Get-AuditPaths -Mode Test
Initialize-AuditFolders $Paths

$Rand = New-Object System.Random($Seed)
$Now = (Get-Date).ToUniversalTime()
$Iso = 'yyyy-MM-ddTHH:mm:ssZ'
$Inv = [Globalization.CultureInfo]::InvariantCulture

$First = 'Alex', 'Blake', 'Casey', 'Dana', 'Elliot', 'Frankie', 'Gray', 'Harper', 'Indy', 'Jordan', 'Kai', 'Logan',
'Morgan', 'Noel', 'Oakley', 'Parker', 'Quinn', 'Riley', 'Sage', 'Taylor', 'Umber', 'Val', 'Wren', 'Avery', 'Rowan'
$Last = 'Ashford', 'Brookes', 'Calder', 'Dunmore', 'Ellison', 'Fairbank', 'Garrow', 'Hollis', 'Ingram', 'Jessop',
'Kendrick', 'Lowell', 'Marsh', 'Northcott', 'Orwin', 'Pemberton', 'Radley', 'Stanton', 'Thorne', 'Upton', 'Varley', 'Whitlock'

$Paid = @('Microsoft 365 Business Premium', 'Microsoft 365 Business Standard', 'Microsoft 365 E3',
    'Exchange Online (Plan 1)', 'Power BI Pro', 'Microsoft 365 Business Basic')
$Free = @('Microsoft Power Automate Free', 'Microsoft Teams Exploratory', 'Microsoft Fabric (Free)')

function Get-Pick($List) { return $List[$Rand.Next($List.Count)] }

function Get-DaysAgoIso([double]$Days) {
    return $Now.AddDays(-$Days).AddMinutes(-$Rand.Next(0, 600)).ToString($Iso, $Inv)
}

$LicenceRows = New-Object System.Collections.Generic.List[object]
$Activity = New-Object System.Collections.Generic.List[object]
$Used = @{}

for ($i = 0; $i -lt $UserCount; $i++) {
    do {
        $F = Get-Pick $First; $L = Get-Pick $Last
        $Key = "$F.$L"
    } while ($Used.ContainsKey($Key))
    $Used[$Key] = $true

    $Id = [guid]::NewGuid().ToString()
    $IsGuest = ($Rand.NextDouble() -lt 0.12)
    if ($IsGuest) {
        $Upn = ("{0}.{1}_partner.example#EXT#@demo.example.com" -f $F, $L).ToLower()
        $Name = "$F $L (Partner)"
    }
    else {
        $Upn = ("{0}.{1}@demo.example.com" -f $F, $L).ToLower()
        $Name = "$F $L"
    }

    # Activity profile
    $R = $Rand.NextDouble()
    if ($R -lt 0.62) { $DaysAgo = $Rand.Next(0, 25) }        # active
    elseif ($R -lt 0.74) { $DaysAgo = $Rand.Next(31, 60) }   # 31-60
    elseif ($R -lt 0.82) { $DaysAgo = $Rand.Next(61, 90) }   # 61-90
    elseif ($R -lt 0.90) { $DaysAgo = $Rand.Next(91, 150) }  # over 90
    else { $DaysAgo = $null }                                # never seen

    # Licences
    $Licences = 'Unlicensed'
    if (-not $IsGuest -and $Rand.NextDouble() -lt 0.88) {
        $Set = @(Get-Pick $Paid)
        if ($Rand.NextDouble() -lt 0.25) { $Set += Get-Pick $Paid }
        if ($Rand.NextDouble() -lt 0.30) { $Set += Get-Pick $Free }
        $Licences = (@($Set | Select-Object -Unique) -join '+')
    }
    elseif ($IsGuest -and $Rand.NextDouble() -lt 0.2) { $Licences = Get-Pick $Free }

    $Blocked = 'False'
    if ($Rand.NextDouble() -lt 0.06) { $Blocked = 'True' }

    $MemberType = 'Member'
    if ($IsGuest) { $MemberType = 'Guest' }

    $LicenceRows.Add([pscustomobject][ordered]@{
            'Display name'        = $Name
            'User principal name' = $Upn
            'Object Id'           = $Id
            'Block credential'    = $Blocked
            'Licenses'            = $Licences
            'Department'          = (Get-Pick @('Finance', 'Operations', 'Sales', 'IT', 'HR', 'Projects'))
            'Usage location'      = 'GB'
            'MemberType'          = $MemberType
        })

    if ($null -ne $DaysAgo) {
        $LastI = Get-DaysAgoIso ($DaysAgo + $Rand.Next(0, 3))
        $LastN = Get-DaysAgoIso $DaysAgo
        if ($Rand.NextDouble() -lt 0.15) { $LastI = '' }   # only background (non-interactive) activity
        $Activity.Add([pscustomobject]@{
                UserId = $Id; Upn = $Upn; Name = $Name
                LastInteractive = $LastI; LastNonInteractive = $LastN
            })
    }
}

# A couple of sign-in identities that are not in the licence export (e.g. deleted accounts)
foreach ($n in 1..2) {
    $Activity.Add([pscustomobject]@{
            UserId = [guid]::NewGuid().ToString(); Upn = "former.user$n@demo.example.com"; Name = "Former User $n"
            LastInteractive = (Get-DaysAgoIso (40 + $n * 20)); LastNonInteractive = ''
        })
}

Write-JsonFile -Path $Paths.ActivityState -Depth 4 -Object ([ordered]@{
        UpdatedUtc      = $Now.ToString($Iso, $Inv)
        HistoryStartUtc = $Now.AddDays(-150).ToString($Iso, $Inv)
        Users           = $Activity.ToArray()
    })

Get-ChildItem -LiteralPath $Paths.Licenses -Filter '*.csv' -File -ErrorAction SilentlyContinue | Remove-Item -Force
$LicenceFile = Join-Path $Paths.Licenses ('Licenses_{0}_demo_active_users.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$LicenceRows | Export-Csv -LiteralPath $LicenceFile -NoTypeInformation -Encoding UTF8

# Run status, checkpoint and a short notification history
$Started = (Get-Date).AddMinutes(-47)
Write-JsonFile -Path $Paths.Checkpoint -Object ([ordered]@{
        LastSuccessfulEndUtc = $Now.AddMinutes(-2).ToString($Iso, $Inv)
        LastWindowStartUtc   = $Now.AddHours(-30).ToString($Iso, $Inv)
        LastRunId            = $Started.ToString('yyyyMMdd_HHmmss')
        LastRunCompletedLocal = (Get-Date).AddMinutes(-2).ToString('yyyy-MM-dd HH:mm:ss')
        Mode                 = 'Test'
        InteractiveNew       = 1184
        NonInteractiveNew    = 39652
    })

$LogFile = Join-Path $Paths.Logs ("Collector_{0}.log" -f $Started.ToString('yyyyMMdd_HHmmss'))
@(
    "[$($Started.ToString('yyyy-MM-dd HH:mm:ss'))] [INFO] DEMO DATA - fictional log for screenshots"
    "[$($Started.ToString('yyyy-MM-dd HH:mm:ss'))] [INFO] Window: 30.0 h in 2 chunk(s) of 24 h, slices of 6 h"
    "[$($Started.AddMinutes(21).ToString('yyyy-MM-dd HH:mm:ss'))] [INFO] Chunk 1/2 committed: interactive +702, non-interactive +23410"
    "[$($Started.AddMinutes(44).ToString('yyyy-MM-dd HH:mm:ss'))] [INFO] Chunk 2/2 committed: interactive +482, non-interactive +16242"
    "[$($Started.AddMinutes(45).ToString('yyyy-MM-dd HH:mm:ss'))] [INFO] Graph pages: 84, retries: 1, slice splits: 0"
    "[$($Started.AddMinutes(45).ToString('yyyy-MM-dd HH:mm:ss'))] [INFO] FINAL STATUS: SUCCESS"
) | Set-Content -LiteralPath $LogFile -Encoding UTF8

if (Test-Path -LiteralPath $Paths.Notifications) { Remove-Item -LiteralPath $Paths.Notifications -Force }
Add-AuditNotification -Paths $Paths -Level 'info' -Title 'Licence file uploaded' -Message "$($LicenceRows.Count) users loaded from demo_active_users.csv"
Add-AuditNotification -Paths $Paths -Level 'warning' -Title 'Update stopped' -Message 'Stopped from the dashboard. Checkpoint unchanged; the next run resumes from the same point.'
Add-AuditNotification -Paths $Paths -Level 'success' -Title 'Entra sign-in collection succeeded (Test)' -Message 'Interactive +1184, non-interactive +39652 new sign-ins in 45 min.'

$Result = Invoke-EntraReportBuild -Paths $Paths

Write-JsonFile -Path $Paths.LastRun -Object ([ordered]@{
        Status        = 'Success'
        RunId         = $Started.ToString('yyyyMMdd_HHmmss')
        Trigger       = 'Scheduled'
        Mode          = 'Test'
        StartedLocal  = $Started.ToString('yyyy-MM-dd HH:mm:ss')
        FinishedLocal = (Get-Date).AddMinutes(-2).ToString('yyyy-MM-dd HH:mm:ss')
        Message       = "Interactive +1184, non-interactive +39652 new sign-ins in 45 min. Licences: $($Result.RemoveCount) to remove, $($Result.ReviewCount) to review."
        LogFile       = $LogFile
    })

Write-Host "Demo data written to $($Paths.DataRoot)"
Write-Host "Users: $($Result.Users)  Remove: $($Result.RemoveCount)  Review: $($Result.ReviewCount)"
Write-Host 'Open it with:  .\Scripts\Start-EntraDashboard.ps1 -Mode Test'

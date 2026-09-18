<#
.SYNOPSIS
    M365admintools.com - Author Charles Arconi - updated 7/7/2026
    Generates a monthly backup session summary report for Veeam Backup & Replication v13.

.DESCRIPTION
    Collects all backup, backup copy, replication, and agent job sessions for a
    given calendar month and produces an HTML report plus optional CSV export.

    Report includes:
      - Executive summary (success rate, totals, data processed)
      - Per-job rollup (runs, success/warning/failure counts, avg duration)
      - Full session detail table
      - Failed and warning sessions called out separately

.PARAMETER Month
    Target month as an integer 1-12. Defaults to the previous calendar month.

.PARAMETER Year
    Target year. Defaults to the year of the previous calendar month.

.PARAMETER OutputPath
    Directory for report output. Defaults to C:\Reports.

.PARAMETER ExportCsv
    Also write the session detail table to CSV.

.PARAMETER Server
    VBR server to connect to. Defaults to localhost.

.EXAMPLE
    .\VBR-Monthly-Report.ps1

.EXAMPLE
    .\VBR-Monthly-Report.ps1 -Month 6 -Year 2026 -OutputPath D:\Reports -ExportCsv

.NOTES
    Requires the Veeam.Backup.PowerShell module (VBR v12+). Run elevated on the
    VBR server, or supply -Server for a remote connection.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 12)]
    [int]$Month,

    [int]$Year,

    [string]$OutputPath = 'C:\Reports',

    [switch]$ExportCsv,

    [string]$Server = 'localhost'
)

# ----------------------------------------------------------------------------
# Setup
# ----------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'

try {
    Import-Module Veeam.Backup.PowerShell -DisableNameChecking
}
catch {
    Write-Error "Could not load Veeam.Backup.PowerShell. Confirm VBR v12+ console is installed. $_"
    exit 1
}

# Connect if not already connected
if (-not (Get-VBRServerSession -ErrorAction SilentlyContinue)) {
    Connect-VBRServer -Server $Server
}

# Resolve reporting window: default to previous calendar month
if (-not $Month -or -not $Year) {
    $lastMonth = (Get-Date).AddMonths(-1)
    if (-not $Month) { $Month = $lastMonth.Month }
    if (-not $Year)  { $Year  = $lastMonth.Year }
}

$startDate = Get-Date -Year $Year -Month $Month -Day 1 -Hour 0 -Minute 0 -Second 0
$endDate   = $startDate.AddMonths(1)
$periodLabel = $startDate.ToString('MMMM yyyy')

Write-Host "Collecting sessions for $periodLabel ..." -ForegroundColor Cyan

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# ----------------------------------------------------------------------------
# Collect sessions
# ----------------------------------------------------------------------------

$allSessions = @()

# Standard backup / backup copy / replication sessions
$vbrSessions = Get-VBRBackupSession | Where-Object {
    $_.EndTime -ge $startDate -and $_.EndTime -lt $endDate
}

foreach ($s in $vbrSessions) {
    $duration = if ($s.EndTime -and $s.CreationTime) {
        ($s.EndTime - $s.CreationTime)
    } else { New-TimeSpan }

    $transferred = 0
    $processed   = 0
    if ($s.Progress) {
        $transferred = $s.Progress.TransferedSize
        $processed   = $s.Progress.ProcessedUsedSize
    }

    $allSessions += [PSCustomObject]@{
        JobName        = $s.JobName
        JobType        = $s.JobType
        Start          = $s.CreationTime
        End            = $s.EndTime
        DurationMin    = [math]::Round($duration.TotalMinutes, 1)
        Result         = $s.Result.ToString()
        ProcessedGB    = [math]::Round($processed / 1GB, 2)
        TransferredGB  = [math]::Round($transferred / 1GB, 2)
    }
}

# Agent job sessions (physical servers / workstations)
try {
    $agentSessions = Get-VBRComputerBackupJobSession | Where-Object {
        $_.EndTime -ge $startDate -and $_.EndTime -lt $endDate
    }

    foreach ($s in $agentSessions) {
        $duration = if ($s.EndTime -and $s.CreationTime) {
            ($s.EndTime - $s.CreationTime)
        } else { New-TimeSpan }

        $allSessions += [PSCustomObject]@{
            JobName       = $s.JobName
            JobType       = 'Agent'
            Start         = $s.CreationTime
            End           = $s.EndTime
            DurationMin   = [math]::Round($duration.TotalMinutes, 1)
            Result        = $s.Result.ToString()
            ProcessedGB   = 0
            TransferredGB = 0
        }
    }
}
catch {
    Write-Warning "Agent session collection skipped: $($_.Exception.Message)"
}

if ($allSessions.Count -eq 0) {
    Write-Warning "No sessions found for $periodLabel. Check retention on session history."
    Disconnect-VBRServer
    exit 0
}

$allSessions = $allSessions | Sort-Object Start

# ----------------------------------------------------------------------------
# Aggregate
# ----------------------------------------------------------------------------

$total    = $allSessions.Count
$success  = ($allSessions | Where-Object Result -eq 'Success').Count
$warning  = ($allSessions | Where-Object Result -eq 'Warning').Count
$failed   = ($allSessions | Where-Object Result -eq 'Failed').Count
$rate     = if ($total) { [math]::Round(($success / $total) * 100, 1) } else { 0 }
$totalGB  = [math]::Round(($allSessions | Measure-Object TransferredGB -Sum).Sum, 2)

$jobRollup = $allSessions | Group-Object JobName | ForEach-Object {
    $g = $_.Group
    [PSCustomObject]@{
        JobName      = $_.Name
        JobType      = ($g | Select-Object -First 1).JobType
        Runs         = $_.Count
        Success      = ($g | Where-Object Result -eq 'Success').Count
        Warning      = ($g | Where-Object Result -eq 'Warning').Count
        Failed       = ($g | Where-Object Result -eq 'Failed').Count
        SuccessRate  = [math]::Round((($g | Where-Object Result -eq 'Success').Count / $_.Count) * 100, 1)
        AvgDurMin    = [math]::Round((($g | Measure-Object DurationMin -Average).Average), 1)
        TransferGB   = [math]::Round((($g | Measure-Object TransferredGB -Sum).Sum), 2)
    }
} | Sort-Object JobName

$problems = $allSessions | Where-Object { $_.Result -ne 'Success' } | Sort-Object Start -Descending

# ----------------------------------------------------------------------------
# Build HTML
# ----------------------------------------------------------------------------

function ConvertTo-HtmlRows {
    param([object[]]$Data, [string[]]$Columns)
    $sb = New-Object System.Text.StringBuilder
    $i = 0
    foreach ($row in $Data) {
        $cls = if ($i % 2 -eq 0) { 'even' } else { 'odd' }
        [void]$sb.Append("<tr class='$cls'>")
        foreach ($c in $Columns) {
            $val = $row.$c
            if ($val -is [datetime]) { $val = $val.ToString('yyyy-MM-dd HH:mm') }
            $cell = "<td>$val</td>"
            if ($c -eq 'Result') {
                switch ($val) {
                    'Success' { $cell = "<td class='ok'>$val</td>" }
                    'Warning' { $cell = "<td class='warn'>$val</td>" }
                    'Failed'  { $cell = "<td class='fail'>$val</td>" }
                }
            }
            [void]$sb.Append($cell)
        }
        [void]$sb.Append('</tr>')
        $i++
    }
    return $sb.ToString()
}

$sessionCols = 'JobName','JobType','Start','End','DurationMin','Result','ProcessedGB','TransferredGB'
$rollupCols  = 'JobName','JobType','Runs','Success','Warning','Failed','SuccessRate','AvgDurMin','TransferGB'

$css = @'
<style>
body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #333; margin: 30px; }
h1 { color: #1F3864; font-size: 20pt; margin-bottom: 4px; }
h2 { color: #1F3864; font-size: 14pt; border-bottom: 2px solid #1F3864; padding-bottom: 4px; margin-top: 28px; }
.subtitle { color: #666; font-size: 10pt; margin-bottom: 20px; }
table { border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 10pt; }
th { background: #1F3864; color: #fff; text-align: left; padding: 7px 9px; }
td { padding: 6px 9px; border-bottom: 1px solid #ddd; }
tr.odd { background: #F2F5FA; }
.ok { color: #107C10; font-weight: bold; }
.warn { color: #B7791F; font-weight: bold; }
.fail { color: #C00000; font-weight: bold; }
.cards { display: flex; gap: 14px; flex-wrap: wrap; margin-top: 14px; }
.card { border: 1px solid #d0d7e2; border-radius: 6px; padding: 14px 20px; min-width: 130px; }
.card .num { font-size: 22pt; font-weight: bold; color: #1F3864; }
.card .lbl { font-size: 9pt; color: #666; text-transform: uppercase; }
.footer { margin-top: 34px; font-size: 9pt; color: #888; border-top: 1px solid #ddd; padding-top: 8px; }
</style>
'@

$html = @"
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Veeam Monthly Backup Report - $periodLabel</title>$css</head>
<body>
<h1>Veeam Monthly Backup Report</h1>
<div class="subtitle">Reporting period: $periodLabel &nbsp;|&nbsp; VBR server: $Server &nbsp;|&nbsp; Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')</div>

<h2>Summary</h2>
<div class="cards">
  <div class="card"><div class="num">$total</div><div class="lbl">Total Sessions</div></div>
  <div class="card"><div class="num">$rate%</div><div class="lbl">Success Rate</div></div>
  <div class="card"><div class="num">$success</div><div class="lbl">Success</div></div>
  <div class="card"><div class="num">$warning</div><div class="lbl">Warning</div></div>
  <div class="card"><div class="num">$failed</div><div class="lbl">Failed</div></div>
  <div class="card"><div class="num">$totalGB</div><div class="lbl">GB Transferred</div></div>
</div>

<h2>Per-Job Rollup</h2>
<table><tr>$(($rollupCols | ForEach-Object { "<th>$_</th>" }) -join '')</tr>
$(ConvertTo-HtmlRows -Data $jobRollup -Columns $rollupCols)
</table>

<h2>Sessions Requiring Attention ($($problems.Count))</h2>
$(if ($problems.Count) {
"<table><tr>$(($sessionCols | ForEach-Object { "<th>$_</th>" }) -join '')</tr>$(ConvertTo-HtmlRows -Data $problems -Columns $sessionCols)</table>"
} else { "<p>No warnings or failures recorded in this period.</p>" })

<h2>All Sessions</h2>
<table><tr>$(($sessionCols | ForEach-Object { "<th>$_</th>" }) -join '')</tr>
$(ConvertTo-HtmlRows -Data $allSessions -Columns $sessionCols)
</table>

<div class="footer">Generated by VBR-Monthly-Report.ps1 &nbsp;|&nbsp; m365admintools.com - Charles Arconi</div>
</body></html>
"@

$stamp    = $startDate.ToString('yyyy-MM')
$htmlFile = Join-Path $OutputPath "VBR-Monthly-Report-$stamp.html"
$html | Out-File -FilePath $htmlFile -Encoding UTF8

Write-Host "HTML report written to $htmlFile" -ForegroundColor Green

if ($ExportCsv) {
    $csvFile = Join-Path $OutputPath "VBR-Monthly-Sessions-$stamp.csv"
    $allSessions | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
    Write-Host "CSV written to $csvFile" -ForegroundColor Green
}

Write-Host ""
Write-Host "$periodLabel : $total sessions, $rate% success, $failed failed, $warning warning" -ForegroundColor Cyan

Disconnect-VBRServer

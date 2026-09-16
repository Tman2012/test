<#
.SYNOPSIS
    Runs a stress test directly on this session host, tracking session start through
    logoff. No Azure Resource Manager calls, no subscription ID, no resource group needed -
    this runs INSIDE the VM being tested.

.DESCRIPTION
    Meant to be triggered automatically at logon (see the scheduled-task registration added
    to fslogix_config.ps1/fslogix_cc_config.ps1), since a single real test account can only
    produce one real logon - so this script simulates the other users on a 3-users-per-VM
    host itself, by fanning out -Sessions concurrent background jobs within that one logon,
    rather than requiring -Sessions separate real logons.

    Per run:
      1. Detects when the CURRENT session actually started, best-effort, via explorer.exe's
         process start time for this session ID (explorer starts once the desktop loads,
         right after profile/FSLogix mount completes - a reasonable proxy for "login done",
         and normally near-simultaneous with this script's own start if launched by the
         at-logon scheduled task).
      2. Pulls any Microsoft-FSLogix-Apps/Operational event log entries since that time, as
         supporting diagnostic data (not asserting specific event ID meanings - just surfaced
         raw for you to read).
      3. Runs -Sessions parallel workload streams for -DurationMinutes each: CPU burn, disk
         I/O, Word/Excel COM automation (Microsoft 365 Apps for enterprise), a Teams process
         launch/stop, and a headless Chrome launch representing the Genesys Cloud agent
         desktop's browser + background-assistant footprint. Genesys Cloud itself is NOT
         logged into - see the Chrome step's own comment for why.
      4. Samples local CPU for the whole window via Get-Counter and evaluates it against
         -CpuThresholdPercent.
      5. Writes an HTML report (uniquely named per run, so a scheduled task firing on every
         logon never overwrites a previous run's report), then - ONLY if -LogoffAtEnd is
         passed - logs off.

    -LogoffAtEnd is OFF by default and must be requested explicitly. Logging off ends the
    ENTIRE Windows session for this user. Only pass it when this run owns its own dedicated
    real logon and nothing else needs that session to stay open.

.PARAMETER DurationMinutes
    How long the workload runs. Defaults to 15.

.PARAMETER Sessions
    Number of concurrent simulated sessions to fan out within this one logon. Defaults to 3
    (matches production's 3-users-per-VM ratio).

.PARAMETER CpuThresholdPercent
    Average CPU above this over the test window fails the test. Defaults to 85.

.PARAMETER ReportPath
    Where to write the HTML report. Defaults to a file under a 'reports' subfolder next to
    this script, uniquely named per computer/user/process/timestamp so repeated runs (e.g.
    firing on every logon) never overwrite each other.

.PARAMETER LogoffAtEnd
    If set, logs off this session after the report is written. OFF by default - see
    .DESCRIPTION for why this is dangerous to set casually.

.EXAMPLE
    # Typical: fires automatically at logon via the registered scheduled task
    ./Start-StressTest.ps1 -DurationMinutes 15 -Sessions 3

.EXAMPLE
    # Real logon, own dedicated session, OK to log off once the test completes
    ./Start-StressTest.ps1 -DurationMinutes 20 -Sessions 3 -LogoffAtEnd

.NOTES
    No Az PowerShell module required. Run as the same user context real sessions run as (not
    an elevated admin RDP session) if you want the Office/Teams/Chrome results to be
    representative.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 240)]
    [int]$DurationMinutes = 15,

    [ValidateRange(1, 10)]
    [int]$Sessions = 3,

    [ValidateRange(1, 100)]
    [int]$CpuThresholdPercent = 85,

    [string]$ReportPath,

    [switch]$LogoffAtEnd
)

$ErrorActionPreference = 'Stop'

$sessionId = (Get-Process -Id $PID).SessionId

if (-not $ReportPath) {
    $reportsDir = Join-Path $PSScriptRoot 'reports'
    New-Item -Path $reportsDir -ItemType Directory -Force | Out-Null
    # Computer + user + PID + timestamp - unique even if two instances start in the same
    # second under the same user/session, since PID never collides between running processes.
    $ReportPath = Join-Path $reportsDir "stress-test-$env:COMPUTERNAME-$env:USERNAME-pid$PID-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
}

function Write-Step {
    param([string]$Text)
    Write-Host "`n=== $Text ===" -ForegroundColor Cyan
}

# ------------------------------------------------------------
# Best-effort real login time for this session - explorer.exe's start time
# for this session ID. If this instance is sharing a session with other
# already-running instances, explorer already started long ago and this
# just reports that, which is still accurate (it IS when this session began).
# ------------------------------------------------------------
function Get-SessionLoginTime {
    param([int]$SessionId)

    $explorerProc = Get-Process -Name explorer -ErrorAction SilentlyContinue |
        Where-Object { $_.SessionId -eq $SessionId } |
        Sort-Object StartTime | Select-Object -First 1

    return $explorerProc.StartTime
}

# ------------------------------------------------------------
# Best-effort FSLogix diagnostic pull - surfaced raw in the report, not
# parsed/asserted, since exact event ID semantics aren't guaranteed here.
# ------------------------------------------------------------
function Get-FSLogixEventsSince {
    param([datetime]$Since)

    try {
        Get-WinEvent -LogName 'Microsoft-FSLogix-Apps/Operational' -ErrorAction Stop |
            Where-Object { $_.TimeCreated -ge $Since } |
            Sort-Object TimeCreated |
            Select-Object TimeCreated, Id, LevelDisplayName, Message
    } catch {
        @()
    }
}

# ------------------------------------------------------------
# Workload for one simulated session - CPU burn + disk I/O + Word/Excel/
# Teams/Chrome. Run as a scriptblock via Start-Job, one job per -Sessions,
# so all simulated sessions run concurrently within this one logon.
# ------------------------------------------------------------
$workload = {
    param($Deadline, $SessionId)

    $tempDir = Join-Path $env:TEMP "stress-session$SessionId"
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

    $chromePath = @(
        (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe')
        (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe')
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    # New Teams (per-user install) and classic Teams (machine-wide) live in different places -
    # check both since we don't know which one this image ended up with.
    $teamsPath = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\ms-teams.exe')
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Teams\current\Teams.exe')
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    $counts = [ordered]@{ Iterations = 0; WordFail = 0; ExcelFail = 0; TeamsFail = 0; ChromeFail = 0 }

    while ((Get-Date) -lt $Deadline) {
        $counts.Iterations++

        # CPU burn - a few hundred ms of tight math, enough to show up in the CPU counter
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $x = 0
        while ($sw.ElapsedMilliseconds -lt 500) { $x = [Math]::Sqrt($x + 1) }

        # Disk I/O - write/read/delete a ~5MB file
        $filePath = Join-Path $tempDir "io_$($counts.Iterations).tmp"
        [System.IO.File]::WriteAllBytes($filePath, (New-Object byte[] (5MB)))
        [void][System.IO.File]::ReadAllBytes($filePath)
        Remove-Item $filePath -ErrorAction SilentlyContinue

        # Word
        $word = $null
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $doc = $word.Documents.Add()
            $doc.Content.Text = ("Stress test session $SessionId iteration $($counts.Iterations). " * 50)
            $docPath = Join-Path $tempDir "doc_$($counts.Iterations).docx"
            $doc.SaveAs([ref]$docPath)
            $doc.Close()
            Remove-Item $docPath -ErrorAction SilentlyContinue
        } catch {
            $counts.WordFail++
        } finally {
            if ($word) { $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null }
        }

        # Excel
        $excel = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $wb = $excel.Workbooks.Add()
            $wb.Sheets.Item(1).Cells.Item(1, 1) = "Stress test session $SessionId iteration $($counts.Iterations)"
            $xlsxPath = Join-Path $tempDir "wb_$($counts.Iterations).xlsx"
            $wb.SaveAs($xlsxPath)
            $wb.Close($false)
            Remove-Item $xlsxPath -ErrorAction SilentlyContinue
        } catch {
            $counts.ExcelFail++
        } finally {
            if ($excel) { $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
        }

        # Teams
        try {
            if (-not $teamsPath) { throw "Teams executable not found" }
            $proc = Start-Process -FilePath $teamsPath -PassThru -ErrorAction Stop
            Start-Sleep -Seconds 5
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        } catch {
            $counts.TeamsFail++
        }

        # Chrome - representative Genesys Cloud agent-desktop footprint (no login/credentials -
        # Genesys Cloud is a SaaS platform reached via browser + its Background Assistant
        # helper; driving a real agent session would need test credentials this script has no
        # business holding)
        try {
            if (-not $chromePath) { throw "chrome.exe not found" }
            $proc = Start-Process -FilePath $chromePath -ArgumentList '--headless=new', '--disable-gpu', 'about:blank' -PassThru -ErrorAction Stop
            Start-Sleep -Seconds 3
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        } catch {
            $counts.ChromeFail++
        }

        Start-Sleep -Seconds 2
    }

    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    [PSCustomObject]@{
        SessionId  = $SessionId
        Iterations = $counts.Iterations
        WordFail   = $counts.WordFail
        ExcelFail  = $counts.ExcelFail
        TeamsFail  = $counts.TeamsFail
        ChromeFail = $counts.ChromeFail
    }
}

# ------------------------------------------------------------
# Local CPU sampling for the test window - no Azure Monitor round-trip needed.
# ------------------------------------------------------------
function Start-CpuSampler {
    param([datetime]$Deadline)

    Start-Job -ScriptBlock {
        param($Deadline)
        $samples = @()
        while ((Get-Date) -lt $Deadline) {
            $value = (Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction SilentlyContinue).CounterSamples.CookedValue
            if ($null -ne $value) { $samples += $value }
            Start-Sleep -Seconds 5
        }
        $samples
    } -ArgumentList $Deadline
}

# ------------------------------------------------------------
# Build the self-contained HTML report (inline CSS, no external assets).
# ------------------------------------------------------------
function New-HtmlReport {
    param(
        [hashtable]$Identity,
        [double]$AvgCpu,
        [double]$MaxCpu,
        [Nullable[bool]]$Pass,
        [array]$SessionResults,
        [array]$FSLogixEvents,
        [hashtable]$TestParams
    )

    function Enc([string]$s) { [System.Net.WebUtility]::HtmlEncode([string]$s) }

    $statusLabel = if ($Pass -eq $true) { 'PASSED' } elseif ($Pass -eq $false) { 'FAILED' } else { 'INCOMPLETE' }
    $statusColor = if ($Pass -eq $true) { '#22c55e' } elseif ($Pass -eq $false) { '#ef4444' } else { '#f59e0b' }
    $cpuText = if ($null -ne $AvgCpu) { "$AvgCpu% avg / $MaxCpu% max" } else { 'No CPU samples collected' }
    $loginText = if ($Identity.LoginTime) { $Identity.LoginTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Unknown (explorer.exe not found for this session)' }
    $logoffText = if ($TestParams.LogoffAtEnd) { $TestParams.EndTime } else { 'Not requested (-LogoffAtEnd not passed) - session left open' }
    $totalIterations = ($SessionResults | Measure-Object -Property Iterations -Sum).Sum
    $totalAppFailures = ($SessionResults | ForEach-Object { $_.WordFail + $_.ExcelFail + $_.TeamsFail + $_.ChromeFail } | Measure-Object -Sum).Sum

    $sessionRows = ($SessionResults | ForEach-Object {
        "<tr><td>Session $($_.SessionId)</td><td>$($_.Iterations)</td><td>$($_.WordFail)</td><td>$($_.ExcelFail)</td><td>$($_.TeamsFail)</td><td>$($_.ChromeFail)</td></tr>"
    }) -join "`n"
    if (-not $sessionRows) { $sessionRows = '<tr><td colspan="6" class="muted">No session data reported</td></tr>' }

    $fslogixRows = ($FSLogixEvents | ForEach-Object {
        "<tr><td>$(Enc($_.TimeCreated.ToString('HH:mm:ss')))</td><td>$($_.Id)</td><td>$(Enc($_.LevelDisplayName))</td><td>$(Enc($_.Message))</td></tr>"
    }) -join "`n"
    if (-not $fslogixRows) { $fslogixRows = '<tr><td colspan="4" class="muted">No FSLogix Operational events found since login (log may not exist on this image, or nothing logged in the window)</td></tr>' }

    @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Stress Test Report - $(Enc($Identity.ComputerName)) pid $($Identity.Pid)</title>
<style>
  :root { color-scheme: light; }
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 40px 20px;
    background: #f4f5f7;
    font-family: -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    color: #1f2430;
  }
  .container { max-width: 820px; margin: 0 auto; }
  .banner {
    border-radius: 12px; padding: 28px 32px; margin-bottom: 24px;
    background: $statusColor; color: #fff;
    box-shadow: 0 4px 14px rgba(0,0,0,0.12);
  }
  .banner h1 { margin: 0 0 6px; font-size: 22px; }
  .banner .sub { opacity: 0.9; font-size: 14px; }
  .cards { display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap; }
  .card {
    flex: 1; min-width: 140px; background: #fff; border-radius: 10px; padding: 16px 18px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
  }
  .card .label { font-size: 12px; text-transform: uppercase; color: #6b7280; letter-spacing: .04em; }
  .card .value { font-size: 20px; font-weight: 600; margin-top: 4px; }
  .panel { background: #fff; border-radius: 10px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); overflow: hidden; margin-bottom: 24px; }
  .panel h2 { font-size: 15px; margin: 0; padding: 16px 20px; border-bottom: 1px solid #eceef1; }
  table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
  th { text-align: left; padding: 10px 20px; background: #fafbfc; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: .03em; }
  td { padding: 8px 20px; border-top: 1px solid #f0f1f3; vertical-align: top; }
  .muted { color: #9ca3af; font-style: italic; }
  .params { font-size: 13px; color: #4b5563; padding: 16px 20px; }
  .params span { display: inline-block; margin-right: 18px; margin-bottom: 6px; }
  .params b { color: #1f2430; }
  .lifecycle { padding: 16px 20px; font-size: 13.5px; }
  .lifecycle div { padding: 6px 0; border-top: 1px solid #f0f1f3; }
  .lifecycle div:first-child { border-top: none; }
  footer { text-align: center; font-size: 12px; color: #9ca3af; margin-top: 24px; }
</style>
</head>
<body>
  <div class="container">
    <div class="banner">
      <h1>AVD Golden Image Stress Test — $statusLabel</h1>
      <div class="sub">$(Enc($Identity.ComputerName)) &middot; $(Enc($Identity.UserName)) &middot; session $($Identity.SessionId) &middot; pid $($Identity.Pid)</div>
    </div>

    <div class="cards">
      <div class="card"><div class="label">CPU (avg / max)</div><div class="value">$cpuText</div></div>
      <div class="card"><div class="label">Sessions</div><div class="value">$($SessionResults.Count)</div></div>
      <div class="card"><div class="label">Iterations</div><div class="value">$totalIterations</div></div>
      <div class="card"><div class="label">App failures</div><div class="value">$totalAppFailures</div></div>
    </div>

    <div class="panel">
      <h2>Session lifecycle</h2>
      <div class="lifecycle">
        <div><b>Login detected</b> $(Enc($loginText))</div>
        <div><b>Test started</b> $(Enc($TestParams.StartTime))</div>
        <div><b>Test ended</b> $(Enc($TestParams.EndTime))</div>
        <div><b>Logoff</b> $(Enc($logoffText))</div>
      </div>
    </div>

    <div class="panel">
      <h2>Test parameters</h2>
      <div class="params">
        <span><b>Duration</b> $($TestParams.DurationMinutes) min</span>
        <span><b>Sessions</b> $($TestParams.Sessions)</span>
        <span><b>CPU threshold</b> $($TestParams.CpuThresholdPercent)%</span>
        <span><b>Report</b> $(Enc($TestParams.ReportPath))</span>
      </div>
    </div>

    <div class="panel">
      <h2>Results by session</h2>
      <table>
        <thead><tr><th>Session</th><th>Iterations</th><th>Word fails</th><th>Excel fails</th><th>Teams fails</th><th>Chrome fails</th></tr></thead>
        <tbody>$sessionRows</tbody>
      </table>
    </div>

    <div class="panel">
      <h2>FSLogix events since login (supporting data, best-effort)</h2>
      <table>
        <thead><tr><th>Time</th><th>Event ID</th><th>Level</th><th>Message</th></tr></thead>
        <tbody>$fslogixRows</tbody>
      </table>
    </div>

    <footer>Generated $(Enc((Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))) by Start-StressTest.ps1</footer>
  </div>
</body>
</html>
"@
}

# ============================================================
# MAIN
# ============================================================
Write-Step "Starting stress test - $env:COMPUTERNAME / $env:USERNAME / session $sessionId ($Sessions simulated sessions, $DurationMinutes min)"

$loginTime = Get-SessionLoginTime -SessionId $sessionId
if ($loginTime) {
    Write-Host "Login detected at $loginTime (explorer.exe start time for this session)"
} else {
    Write-Host "Could not detect login time (no explorer.exe found for session $sessionId)" -ForegroundColor Yellow
}

$startTime = Get-Date
$deadline = $startTime.AddMinutes($DurationMinutes)

$cpuJob = Start-CpuSampler -Deadline $deadline
$workloadJobs = 1..$Sessions | ForEach-Object { Start-Job -ScriptBlock $workload -ArgumentList $deadline, $_ }

$workloadJobs | Wait-Job | Out-Null
$sessionResults = @($workloadJobs | Receive-Job)
$workloadJobs | Remove-Job

$cpuJob | Wait-Job | Out-Null
$cpuSamples = @($cpuJob | Receive-Job)
$cpuJob | Remove-Job
$endTime = Get-Date

$avgCpu = if ($cpuSamples.Count -gt 0) { [math]::Round(($cpuSamples | Measure-Object -Average).Average, 1) } else { $null }
$maxCpu = if ($cpuSamples.Count -gt 0) { [math]::Round(($cpuSamples | Measure-Object -Maximum).Maximum, 1) } else { $null }
$pass = if ($null -ne $avgCpu) { $avgCpu -le $CpuThresholdPercent } else { $null }

$fslogixEvents = Get-FSLogixEventsSince -Since $(if ($loginTime) { $loginTime } else { $startTime })

Write-Step 'Results'
$sessionResults | ForEach-Object {
    Write-Host "  Session $($_.SessionId): $($_.Iterations) iterations, Word fails: $($_.WordFail), Excel fails: $($_.ExcelFail), Teams fails: $($_.TeamsFail), Chrome fails: $($_.ChromeFail)"
}
if ($null -eq $pass) {
    Write-Host "  CPU: UNKNOWN - no samples collected" -ForegroundColor Yellow
} elseif ($pass) {
    Write-Host "  CPU: PASS (avg $avgCpu%, max $maxCpu%)" -ForegroundColor Green
} else {
    Write-Host "  CPU: FAIL (avg $avgCpu%, max $maxCpu% > threshold $CpuThresholdPercent%)" -ForegroundColor Red
}

$identity = @{
    ComputerName = $env:COMPUTERNAME
    UserName     = $env:USERNAME
    SessionId    = $sessionId
    Pid          = $PID
    LoginTime    = $loginTime
}

$testParams = @{
    DurationMinutes     = $DurationMinutes
    Sessions            = $Sessions
    CpuThresholdPercent = $CpuThresholdPercent
    StartTime           = $startTime.ToString('yyyy-MM-dd HH:mm:ss')
    EndTime             = $endTime.ToString('yyyy-MM-dd HH:mm:ss')
    ReportPath          = $ReportPath
    LogoffAtEnd         = [bool]$LogoffAtEnd
}

$html = New-HtmlReport -Identity $identity -AvgCpu $avgCpu -MaxCpu $maxCpu -Pass $pass `
    -SessionResults $sessionResults -FSLogixEvents $fslogixEvents -TestParams $testParams
Set-Content -Path $ReportPath -Value $html -Encoding UTF8
Write-Step "Report written to $ReportPath"

if ($LogoffAtEnd) {
    Write-Host "`n-LogoffAtEnd was passed - logging off session $sessionId in 5 seconds..." -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    logoff $sessionId
} else {
    Write-Host "`nSession left open (-LogoffAtEnd not passed)."
}

if ($pass -eq $false) {
    throw "CPU exceeded threshold (avg $avgCpu% > $CpuThresholdPercent%). See $ReportPath for details."
}

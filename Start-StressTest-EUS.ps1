<#
.SYNOPSIS
    Runs a stress test against existing session host VMs in vdhp-eus-vee-lab by pushing
    concurrent per-session load (stress_workload.ps1) to each target VM via
    Invoke-AzVMRunCommand, evaluates CPU against a threshold, and writes an HTML report.

.DESCRIPTION
    Does NOT create or delete any VMs - point this at hosts already created by
    New-AVDSessionHosts-EUS.ps1 (or any existing host in vm-01/vm-02). For each target VM,
    pushes stress_workload.ps1 via Invoke-AzVMRunCommand, which spawns -SessionsPerVM
    concurrent background jobs on that VM (CPU burn + disk I/O + Word/Excel/Chrome load per
    session, matching production's 3-users-per-VM ratio) for -DurationMinutes. All target VMs
    run their RunCommand concurrently (-AsJob), so a multi-VM test doesn't run serially.

    After the run, pulls Percentage CPU from Get-AzMetric for the test window on each VM,
    combines it with the per-session results each VM reported, and writes a self-contained
    HTML report (no external CSS/JS - safe to open offline) to -ReportPath.

.PARAMETER VMResourceGroup
    Which blue/green VM resource group to target: 'vm-01' or 'vm-02'.

.PARAMETER Numbers
    Specific host numbers to target (e.g. 1,2 -> az-eus-vee-001, -002). Mutually exclusive
    with -All.

.PARAMETER All
    Target every session host VM found in the target resource group. Mutually exclusive with
    -Numbers.

.PARAMETER DurationMinutes
    How long the load runs on each VM. Defaults to 15.

.PARAMETER SessionsPerVM
    Concurrent simulated sessions per VM. Defaults to 3 (matches production).

.PARAMETER CpuThresholdPercent
    Average Percentage CPU above this over the test window fails that VM. Defaults to 85.

.PARAMETER ReportPath
    Where to write the HTML report. Defaults to a timestamped file next to this script, under
    a 'reports' subfolder.

.EXAMPLE
    ./Start-StressTest-EUS.ps1 -VMResourceGroup vm-01 -Numbers 1,2 -DurationMinutes 20

.EXAMPLE
    ./Start-StressTest-EUS.ps1 -VMResourceGroup vm-01 -All -SessionsPerVM 3 -CpuThresholdPercent 90

.NOTES
    Requires an active Az PowerShell context (Connect-AzAccount / Set-AzContext) in the
    calling session - this script does not authenticate on its own.

    Not yet run end-to-end - in particular, whether Word/Excel COM automation comes up at all
    under Invoke-AzVMRunCommand's SYSTEM/no-desktop execution context is unverified (see
    stress_workload.ps1's own caveat). CPU + disk load runs regardless of that, so the test
    still produces real signal even if every app-level step fails on a given host.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('vm-01', 'vm-02')]
    [string]$VMResourceGroup,

    [int[]]$Numbers,
    [switch]$All,

    [ValidateRange(1, 240)]
    [int]$DurationMinutes = 15,

    [ValidateRange(1, 10)]
    [int]$SessionsPerVM = 3,

    [ValidateRange(1, 100)]
    [int]$CpuThresholdPercent = 85,

    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'

if (-not $All -and -not $Numbers) {
    throw "Specify either -Numbers <int[]> or -All."
}
if ($All -and $Numbers) {
    throw "Specify -Numbers or -All, not both."
}

# ============================================================
# POOL CONSTANTS - vdhp-eus-vee-lab (East US)
# Kept in sync with New-AVDSessionHosts-EUS.ps1 / Remove-AVDSessionHosts-EUS.ps1.
# ============================================================
$Pool = @{
    SubscriptionId = '003bcd9a-df27-4cb9-a101-aea65a07750e'
    VMSuffix       = 'eus-vee'
}

$RGDefaults = @{
    'vm-01' = @{ ResourceGroup = 'rg-eus-vee-lab-vm-01' }
    'vm-02' = @{ ResourceGroup = 'rg-eus-vee-lab-vm-02' }
}

$targetRG = $RGDefaults[$VMResourceGroup].ResourceGroup
$WorkloadScriptPath = Join-Path $PSScriptRoot 'stress_workload.ps1'

if (-not $ReportPath) {
    $reportsDir = Join-Path $PSScriptRoot 'reports'
    New-Item -Path $reportsDir -ItemType Directory -Force | Out-Null
    $ReportPath = Join-Path $reportsDir "stress-test-eus-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
}

function Write-Step {
    param([string]$Text)
    Write-Host "`n=== $Text ===" -ForegroundColor Cyan
}

# ------------------------------------------------------------
# Resolve which VM names to target.
# ------------------------------------------------------------
function Resolve-TargetVMNames {
    if (-not $All) {
        return $Numbers | ForEach-Object { "az-$($Pool.VMSuffix)-$('{0:D3}' -f $_)" }
    }

    Write-Step "Scanning $targetRG for session host VMs"
    $pattern = "^az-$($Pool.VMSuffix)-\d{3}$"
    $names = Get-AzVM -ResourceGroupName $targetRG | Select-Object -ExpandProperty Name |
        Where-Object { $_ -match $pattern }
    Write-Host "Found: $($names -join ', ')"
    return $names
}

# ------------------------------------------------------------
# Push the workload to one VM. Never throws - failures are reported in the
# summary so one bad host doesn't stop the rest.
# ------------------------------------------------------------
function Start-VMWorkload {
    param([string]$VmName)

    Write-Host "Starting workload on $VmName ($SessionsPerVM sessions, $DurationMinutes min)..."
    $job = Invoke-AzVMRunCommand -ResourceGroupName $targetRG -VMName $VmName `
        -CommandId 'RunPowerShellScript' -ScriptPath $WorkloadScriptPath `
        -Parameter @{ DurationMinutes = $DurationMinutes; Sessions = $SessionsPerVM } `
        -AsJob

    return @{ Name = $VmName; Job = $job }
}

# ------------------------------------------------------------
# Pull the per-session JSON line (see stress_workload.ps1's "RESULT_JSON:"
# output) out of a RunCommand result's stdout.
# ------------------------------------------------------------
function Get-SessionResultsFromOutput {
    param($RunCommandResult)

    $text = ($RunCommandResult.Value | Select-Object -ExpandProperty Message) -join "`n"
    foreach ($line in $text -split "`r?`n") {
        Write-Host $line
    }

    $jsonLine = ($text -split "`r?`n") | Where-Object { $_ -like 'RESULT_JSON:*' } | Select-Object -Last 1
    if (-not $jsonLine) { return @() }
    try {
        return @(($jsonLine -replace '^RESULT_JSON:', '') | ConvertFrom-Json)
    } catch {
        return @()
    }
}

# ------------------------------------------------------------
# Pull Percentage CPU for one VM over the test window and evaluate it against
# -CpuThresholdPercent.
# ------------------------------------------------------------
function Get-VMCpuResult {
    param(
        [string]$VmName,
        [datetime]$StartTime,
        [datetime]$EndTime
    )

    $vmId = "/subscriptions/$($Pool.SubscriptionId)/resourceGroups/$targetRG/providers/Microsoft.Compute/virtualMachines/$VmName"

    try {
        $metric = Get-AzMetric -ResourceId $vmId -MetricName 'Percentage CPU' `
            -StartTime $StartTime -EndTime $EndTime -TimeGrain '00:01:00' `
            -AggregationType Average -WarningAction SilentlyContinue

        $values = $metric.Data | Where-Object { $null -ne $_.Average } | Select-Object -ExpandProperty Average
        if (-not $values) {
            return @{ AvgCpu = $null; MaxCpu = $null; Pass = $null; Note = 'No metric data returned yet (may need a few minutes to land)' }
        }

        $avg = ($values | Measure-Object -Average).Average
        $max = ($values | Measure-Object -Maximum).Maximum
        return @{ AvgCpu = [math]::Round($avg, 1); MaxCpu = [math]::Round($max, 1); Pass = ($avg -le $CpuThresholdPercent); Note = $null }
    } catch {
        return @{ AvgCpu = $null; MaxCpu = $null; Pass = $null; Note = "Metric query failed: $($_.Exception.Message)" }
    }
}

# ------------------------------------------------------------
# Build the self-contained HTML report (inline CSS, no external assets).
# ------------------------------------------------------------
function New-HtmlReport {
    param(
        [array]$VmResults,
        [hashtable]$TestParams
    )

    $passCount = ($VmResults | Where-Object { $_.Pass -eq $true }).Count
    $failCount = ($VmResults | Where-Object { $_.Pass -eq $false }).Count
    $unknownCount = ($VmResults | Where-Object { $null -eq $_.Pass }).Count
    $overallPass = $failCount -eq 0 -and $unknownCount -eq 0
    $overallLabel = if ($failCount -gt 0) { 'FAILED' } elseif ($unknownCount -gt 0) { 'INCOMPLETE' } else { 'PASSED' }
    $overallColor = if ($failCount -gt 0) { '#ef4444' } elseif ($unknownCount -gt 0) { '#f59e0b' } else { '#22c55e' }

    function Enc([string]$s) { [System.Net.WebUtility]::HtmlEncode($s) }

    $vmRows = ($VmResults | ForEach-Object {
        $statusColor = if ($_.Pass -eq $true) { '#22c55e' } elseif ($_.Pass -eq $false) { '#ef4444' } else { '#f59e0b' }
        $statusLabel = if ($_.Pass -eq $true) { 'PASS' } elseif ($_.Pass -eq $false) { 'FAIL' } else { 'UNKNOWN' }
        $cpuCell = if ($null -ne $_.AvgCpu) { "$($_.AvgCpu)% avg / $($_.MaxCpu)% max" } else { Enc($_.Note) }

        $sessionRows = ($_.Sessions | ForEach-Object {
            "<tr><td>Session $($_.SessionId)</td><td>$($_.Iterations)</td><td>$($_.WordFail)</td><td>$($_.ExcelFail)</td><td>$($_.TeamsFail)</td><td>$($_.ChromeFail)</td></tr>"
        }) -join "`n"
        if (-not $sessionRows) { $sessionRows = '<tr><td colspan="6" class="muted">No session data reported</td></tr>' }

        @"
        <tr class="vm-row">
          <td><strong>$(Enc($_.Name))</strong></td>
          <td>$cpuCell</td>
          <td><span class="badge" style="background:$statusColor">$statusLabel</span></td>
        </tr>
        <tr class="detail-row">
          <td colspan="3">
            <table class="session-table">
              <thead><tr><th>Session</th><th>Iterations</th><th>Word fails</th><th>Excel fails</th><th>Teams fails</th><th>Chrome fails</th></tr></thead>
              <tbody>$sessionRows</tbody>
            </table>
          </td>
        </tr>
"@
    }) -join "`n"

    @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>AVD Stress Test Report - $(Enc($TestParams.VMResourceGroup))</title>
<style>
  :root { color-scheme: light; }
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 40px 20px;
    background: #f4f5f7;
    font-family: -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    color: #1f2430;
  }
  .container { max-width: 900px; margin: 0 auto; }
  .banner {
    border-radius: 12px; padding: 28px 32px; margin-bottom: 24px;
    background: $overallColor; color: #fff;
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
  .card .value { font-size: 26px; font-weight: 600; margin-top: 4px; }
  .panel { background: #fff; border-radius: 10px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); overflow: hidden; margin-bottom: 24px; }
  .panel h2 { font-size: 15px; margin: 0; padding: 16px 20px; border-bottom: 1px solid #eceef1; }
  table { width: 100%; border-collapse: collapse; font-size: 14px; }
  th { text-align: left; padding: 10px 20px; background: #fafbfc; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: .03em; }
  td { padding: 10px 20px; border-top: 1px solid #f0f1f3; }
  .vm-row td { padding-top: 16px; }
  .detail-row td { padding: 0 20px 16px; border-top: none; }
  .session-table { background: #fafbfc; border-radius: 8px; }
  .session-table th, .session-table td { padding: 6px 14px; font-size: 12.5px; }
  .badge { color: #fff; padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; letter-spacing: .03em; }
  .muted { color: #9ca3af; font-style: italic; }
  .params { font-size: 13px; color: #4b5563; padding: 16px 20px; }
  .params span { display: inline-block; margin-right: 18px; }
  .params b { color: #1f2430; }
  footer { text-align: center; font-size: 12px; color: #9ca3af; margin-top: 24px; }
</style>
</head>
<body>
  <div class="container">
    <div class="banner">
      <h1>AVD Golden Image Stress Test — $overallLabel</h1>
      <div class="sub">$(Enc($TestParams.VMResourceGroup)) &middot; $(Enc($TestParams.StartTime)) &rarr; $(Enc($TestParams.EndTime))</div>
    </div>

    <div class="cards">
      <div class="card"><div class="label">VMs tested</div><div class="value">$($VmResults.Count)</div></div>
      <div class="card"><div class="label">Passed</div><div class="value" style="color:#22c55e">$passCount</div></div>
      <div class="card"><div class="label">Failed</div><div class="value" style="color:#ef4444">$failCount</div></div>
      <div class="card"><div class="label">Unknown</div><div class="value" style="color:#f59e0b">$unknownCount</div></div>
    </div>

    <div class="panel">
      <h2>Test parameters</h2>
      <div class="params">
        <span><b>Duration</b> $($TestParams.DurationMinutes) min</span>
        <span><b>Sessions/VM</b> $($TestParams.SessionsPerVM)</span>
        <span><b>CPU threshold</b> $($TestParams.CpuThresholdPercent)%</span>
        <span><b>Resource group</b> $(Enc($TestParams.VMResourceGroup))</span>
      </div>
    </div>

    <div class="panel">
      <h2>Results by VM</h2>
      <table>
        <thead><tr><th>VM</th><th>CPU (avg / max)</th><th>Status</th></tr></thead>
        <tbody>
          $vmRows
        </tbody>
      </table>
    </div>

    <footer>Generated $(Enc((Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))) by Start-StressTest-EUS.ps1</footer>
  </div>
</body>
</html>
"@
}

# ============================================================
# MAIN
# ============================================================
$vmNames = Resolve-TargetVMNames
if (-not $vmNames -or $vmNames.Count -eq 0) {
    Write-Host 'No target VMs found.' -ForegroundColor Yellow
    return
}

Write-Step "Launching workload on $($vmNames.Count) VM(s): $($vmNames -join ', ')"
$startTime = Get-Date
$runs = $vmNames | ForEach-Object { Start-VMWorkload -VmName $_ }

$jobs = $runs.Job
$jobs | Wait-Job | Out-Null
$sessionResultsByVm = @{}
foreach ($r in $runs) {
    Write-Step "Output from $($r.Name)"
    $result = Receive-Job -Job $r.Job
    $sessionResultsByVm[$r.Name] = Get-SessionResultsFromOutput -RunCommandResult $result
}
$jobs | Remove-Job
$endTime = Get-Date

Write-Step "Evaluating CPU against threshold ($CpuThresholdPercent%)"
# Metrics can lag a couple minutes behind real time - give Azure Monitor a moment to catch up.
Start-Sleep -Seconds 90

$vmResults = $vmNames | ForEach-Object {
    $cpu = Get-VMCpuResult -VmName $_ -StartTime $startTime -EndTime $endTime
    [PSCustomObject]@{
        Name     = $_
        AvgCpu   = $cpu.AvgCpu
        MaxCpu   = $cpu.MaxCpu
        Pass     = $cpu.Pass
        Note     = $cpu.Note
        Sessions = $sessionResultsByVm[$_]
    }
}

$vmResults | ForEach-Object {
    if ($null -eq $_.Pass) {
        Write-Host "  $($_.Name): UNKNOWN - $($_.Note)" -ForegroundColor Yellow
    } elseif ($_.Pass) {
        Write-Host "  $($_.Name): PASS (avg $($_.AvgCpu)%, max $($_.MaxCpu)%)" -ForegroundColor Green
    } else {
        Write-Host "  $($_.Name): FAIL (avg $($_.AvgCpu)%, max $($_.MaxCpu)% > threshold $CpuThresholdPercent%)" -ForegroundColor Red
    }
}

$testParams = @{
    VMResourceGroup     = $VMResourceGroup
    DurationMinutes     = $DurationMinutes
    SessionsPerVM       = $SessionsPerVM
    CpuThresholdPercent = $CpuThresholdPercent
    StartTime           = $startTime.ToString('yyyy-MM-dd HH:mm:ss')
    EndTime             = $endTime.ToString('yyyy-MM-dd HH:mm:ss')
}

$html = New-HtmlReport -VmResults $vmResults -TestParams $testParams
Set-Content -Path $ReportPath -Value $html -Encoding UTF8
Write-Step "Report written to $ReportPath"

$failed = $vmResults | Where-Object { $_.Pass -eq $false }
if ($failed) {
    throw "$($failed.Count) VM(s) exceeded the CPU threshold: $($failed.Name -join ', '). See $ReportPath for details."
}

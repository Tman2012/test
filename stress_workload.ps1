<#
.SYNOPSIS
    Runs ON a session host (via Invoke-AzVMRunCommand) to generate concurrent per-session
    load for a fixed duration. Not meant to be run by hand from a workstation.

.DESCRIPTION
    Spawns $Sessions parallel background jobs, each looping until the deadline doing: a short
    CPU burn, a disk write/read/delete cycle, best-effort Word and Excel document create/save/
    close via COM automation (Microsoft 365 Apps for enterprise, confirmed on the golden
    image), a Teams process launch/stop, and a headless Chrome launch representing the Genesys
    Cloud agent desktop's browser + background-assistant footprint.

    Genesys Cloud itself is NOT logged into - it's a SaaS contact-center platform reached
    through a browser tab plus its local "Background Assistant" helper, and driving a real
    agent session would need test credentials this script has no business holding. The
    headless Chrome launch stands in for that footprint (browser process + background helper
    present) without touching real Genesys Cloud auth.

    Word/Excel COM automation and the Teams launch are wrapped in try/catch and do not fail
    the run if they error - launched through RunCommand, they execute as SYSTEM with no
    interactive desktop session, which Office COM and Electron apps like Teams are both known
    to be unreliable under (missing desktop heap, no user profile, first-run prompts). Headless
    Chrome doesn't have that problem (it's designed to run without a desktop), so treat its
    failures as more meaningful than Word/Excel/Teams'. CPU + disk load runs regardless, so the
    test always produces real resource pressure even if every app-level step fails on a given
    host.

.PARAMETER DurationMinutes
    How long each simulated session keeps looping.

.PARAMETER Sessions
    Number of concurrent simulated sessions on this VM (matches production's 3 users/VM).
#>
param(
    [int]$DurationMinutes = 15,
    [int]$Sessions = 3
)

$deadline = (Get-Date).AddMinutes($DurationMinutes)

$workload = {
    param($Deadline, $SessionId)

    $tempDir = Join-Path $env:TEMP "stress-$SessionId"
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

        # CPU burn - a few hundred ms of tight math, enough to show up in Percentage CPU
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $x = 0
        while ($sw.ElapsedMilliseconds -lt 500) { $x = [Math]::Sqrt($x + 1) }

        # Disk I/O - write/read/delete a ~5MB file
        $filePath = Join-Path $tempDir "io_$($counts.Iterations).tmp"
        [System.IO.File]::WriteAllBytes($filePath, (New-Object byte[] (5MB)))
        [void][System.IO.File]::ReadAllBytes($filePath)
        Remove-Item $filePath -ErrorAction SilentlyContinue

        # Word - best-effort, see .DESCRIPTION caveat above
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

        # Excel - best-effort, same caveat as Word
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

        # Teams - best-effort, same SYSTEM/no-desktop caveat as Word/Excel
        try {
            if (-not $teamsPath) { throw "Teams executable not found" }
            $proc = Start-Process -FilePath $teamsPath -PassThru -ErrorAction Stop
            Start-Sleep -Seconds 5
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        } catch {
            $counts.TeamsFail++
        }

        # Chrome - representative Genesys Cloud agent-desktop footprint (no login/credentials)
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

$jobs = 1..$Sessions | ForEach-Object { Start-Job -ScriptBlock $workload -ArgumentList $deadline, $_ }
$jobs | Wait-Job | Out-Null
$results = $jobs | Receive-Job
$jobs | Remove-Job

$results | ForEach-Object {
    Write-Output "Session $($_.SessionId): $($_.Iterations) iterations, Word fails: $($_.WordFail), Excel fails: $($_.ExcelFail), Teams fails: $($_.TeamsFail), Chrome fails: $($_.ChromeFail)"
}
# Machine-readable line for Start-StressTest-EUS.ps1 to parse out of the RunCommand output stream.
Write-Output "RESULT_JSON:$(@($results) | ConvertTo-Json -Compress -Depth 5)"

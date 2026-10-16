# Setup scheduled tasks for NDC, F&F, and ELCM pipelines

param(
    [switch]$Uninstall
)

# ==== Resolve paths ====
$APP_DIR     = Split-Path -Parent $PSScriptRoot        # app/scheduler/ -> app/
$ROOT        = Split-Path -Parent $APP_DIR             # app/ -> project root
$PYTHON      = Join-Path $ROOT ".venv\Scripts\python.exe"
$PYTHONW     = Join-Path $ROOT ".venv\Scripts\pythonw.exe"
$NDC_MODULE  = "app.scripts.automation_scripts.download_ndc_report"
$FNF_MODULE  = "app.scripts.processing_scripts.process_fnf_closed_reports"
$ELCM_MODULE = "app.scripts.automation_scripts.elcm_report_downloader"

# Helper function to dynamically uninstall tasks matching a pattern
function Remove-PipelineTasks {
    param(
        [string]$Pattern,
        [switch]$ShowSkip
    )
    
    $foundTasks = Get-ScheduledTask -TaskName $Pattern -ErrorAction SilentlyContinue
    if ($foundTasks) {
        foreach ($t in $foundTasks) {
            $name = $t.TaskName
            try {
                if ($t.State -eq "Running") {
                    Stop-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
                }
                Unregister-ScheduledTask -TaskName $name -Confirm:$false -ErrorAction Stop
                Write-Host "  [CLEAN] Removed: $name" -ForegroundColor Gray
            } catch {
                Write-Host "  [FAIL]  $name : $_" -ForegroundColor Red
            }
        }
    } elseif ($ShowSkip) {
        Write-Host "  [INFO]  No active tasks matching '$Pattern' found." -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "+------------------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host "|         NDC, F&F, & ELCM AUTOMATION - TASK SCHEDULER                   |" -ForegroundColor Cyan
Write-Host "+------------------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host "|  Project Root : $ROOT" -ForegroundColor Gray
Write-Host "|  Interpreter  : $PYTHONW" -ForegroundColor Gray
Write-Host "|  NDC Module   : $NDC_MODULE" -ForegroundColor Gray
Write-Host "|  F&F Module   : $FNF_MODULE" -ForegroundColor Gray
Write-Host "|  ELCM Module  : $ELCM_MODULE" -ForegroundColor Gray
Write-Host "+------------------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

# ==== Uninstall mode ====
if ($Uninstall) {
    Write-Host "[ACTION] Uninstalling all automation pipeline tasks..." -ForegroundColor Yellow
    Remove-PipelineTasks -Pattern "*_Pipeline_*" -ShowSkip
    Write-Host "`n[DONE] All tasks uninstalled successfully.`n" -ForegroundColor Green
    exit 0
}

# ==== Check virtual environment exists ====
if (-not (Test-Path $PYTHONW) -or -not (Test-Path $PYTHON)) {
    Write-Host "[ERROR] Python virtual environment not found in: $ROOT\.venv" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Please set up the virtual environment first:" -ForegroundColor Yellow
    Write-Host "    cd `"$ROOT`""
    Write-Host "    python -m venv .venv"
    Write-Host "    .venv\Scripts\pip install -e ."
    Write-Host ""
    exit 1
}

# Automatically clean up any existing tasks matching the pattern first
Write-Host "[1/3] Cleaning up previous scheduled tasks..." -ForegroundColor Yellow
Remove-PipelineTasks -Pattern "*_Pipeline_*"

# ==== Query Centralized Schedule Timings from app.core.config ====
Write-Host "`n[2/3] Reading schedule timings from app.core.config..." -ForegroundColor Cyan

$ndcTimesRaw  = & $PYTHON -c "from app.core.config import NDC_SCHEDULE_TIMES; print(chr(44).join(NDC_SCHEDULE_TIMES))"
$fnfTimesRaw  = & $PYTHON -c "from app.core.config import FNF_SCHEDULE_TIMES; print(chr(44).join(FNF_SCHEDULE_TIMES))"
$elcmTimesRaw = & $PYTHON -c "from app.core.config import ELCM_SCHEDULE_TIMES; print(chr(44).join(ELCM_SCHEDULE_TIMES))"

if (-not $ndcTimesRaw) {
    $ndcTimesRaw = "10:00 AM,1:00 PM,3:00 PM,5:00 PM,7:00 PM"
}
if (-not $fnfTimesRaw) {
    $fnfTimesRaw = "11:00 AM,4:00 PM"
}
if (-not $elcmTimesRaw) {
    $elcmTimesRaw = "10:30 AM"
}

# Helper to normalize 12-hour (e.g. "1:00 PM") or 24-hour (e.g. "13:00") into both standard trigger time and clean task name
function Convert-ScheduleTime {
    param([string]$RawTime)

    $trimmed = $RawTime.Trim()
    $parsedDate = $null

    # Normalize dot separators between digits (e.g. "10.00" -> "10:00", "10.30 AM" -> "10:30 AM")
    $normalized = $trimmed -replace '(\d{1,2})\.(\d{2})', '$1:$2'

    # Try 12-hour AM/PM formats first
    $formats12 = @("h:mm tt", "hh:mm tt", "h tt", "hh tt", "h:mmtt", "hh:mmtt")
    foreach ($fmt in $formats12) {
        try {
            $parsedDate = [datetime]::ParseExact($normalized, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
            break
        } catch { }
    }

    # Try 24-hour formats if not matched
    if (-not $parsedDate) {
        $formats24 = @("H:mm", "HH:mm")
        foreach ($fmt in $formats24) {
            try {
                $parsedDate = [datetime]::ParseExact($normalized, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
                break
            } catch { }
        }
    }

    # Fallback to general DateTime parse
    if (-not $parsedDate) {
        try {
            $parsedDate = [datetime]::Parse($normalized, [System.Globalization.CultureInfo]::InvariantCulture)
        } catch {
            $parsedDate = [datetime]::Parse($trimmed)
        }
    }

    $time24 = $parsedDate.ToString("HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
    $cleanTag = $parsedDate.ToString("HHmm", [System.Globalization.CultureInfo]::InvariantCulture)
    $display12 = $parsedDate.ToString("h:mm tt", [System.Globalization.CultureInfo]::InvariantCulture)
    $triggerDate = Get-Date -Hour $parsedDate.Hour -Minute $parsedDate.Minute -Second 0

    return [PSCustomObject]@{
        TriggerDateTime = $triggerDate
        TriggerTime     = $time24
        CleanTag        = $cleanTag
        DisplayTime     = $display12
    }
}

$ndcScheduleList  = $ndcTimesRaw.Trim().Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }
$fnfScheduleList  = $fnfTimesRaw.Trim().Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }
$elcmScheduleList = $elcmTimesRaw.Trim().Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }

# Build task arrays
$ndcTasks = @()
foreach ($timeStr in $ndcScheduleList) {
    $converted = Convert-ScheduleTime -RawTime $timeStr
    $ndcTasks += @{
        Pipeline        = "NDC Report"
        Name            = "NDC_Pipeline_$($converted.CleanTag)"
        Time            = $converted.TriggerTime
        TriggerDateTime = $converted.TriggerDateTime
        DisplayTime     = $converted.DisplayTime
    }
}

$fnfTasks = @()
foreach ($timeStr in $fnfScheduleList) {
    $converted = Convert-ScheduleTime -RawTime $timeStr
    $fnfTasks += @{
        Pipeline        = "F&F Closed"
        Name            = "FNF_Pipeline_$($converted.CleanTag)"
        Time            = $converted.TriggerTime
        TriggerDateTime = $converted.TriggerDateTime
        DisplayTime     = $converted.DisplayTime
    }
}

$elcmTasks = @()
foreach ($timeStr in $elcmScheduleList) {
    $converted = Convert-ScheduleTime -RawTime $timeStr
    $elcmTasks += @{
        Pipeline        = "ELCM Report"
        Name            = "ELCM_Pipeline_$($converted.CleanTag)"
        Time            = $converted.TriggerTime
        TriggerDateTime = $converted.TriggerDateTime
        DisplayTime     = $converted.DisplayTime
    }
}

# ==== Task settings ====
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -WakeToRun `
    -RestartCount             3 `
    -RestartInterval          (New-TimeSpan -Minutes 5) `
    -MultipleInstances        IgnoreNew `
    -ExecutionTimeLimit       (New-TimeSpan -Hours 2) `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries

$ndcAction = New-ScheduledTaskAction `
    -Execute   $PYTHONW `
    -Argument  "-m $NDC_MODULE" `
    -WorkingDirectory $ROOT

$fnfAction = New-ScheduledTaskAction `
    -Execute   $PYTHONW `
    -Argument  "-m $FNF_MODULE" `
    -WorkingDirectory $ROOT

$elcmAction = New-ScheduledTaskAction `
    -Execute   $PYTHONW `
    -Argument  "-m $ELCM_MODULE" `
    -WorkingDirectory $ROOT

# ==== Register tasks ====
$failed = 0
$registeredTasks = @()

Write-Host "`n[3/3] Registering tasks into Windows Task Scheduler..." -ForegroundColor Cyan

foreach ($t in $ndcTasks) {
    try {
        $trigger = New-ScheduledTaskTrigger -Daily -At $t.TriggerDateTime
        Register-ScheduledTask `
            -TaskName $t.Name `
            -Action   $ndcAction `
            -Trigger  $trigger `
            -Settings $settings `
            -RunLevel Limited `
            -Force -ErrorAction Stop | Out-Null
        
        $registeredTasks += [PSCustomObject]@{
            Pipeline  = $t.Pipeline
            TaskName  = $t.Name
            Schedule  = "$($t.DisplayTime) Daily"
            CatchUp   = "Yes (On Boot)"
            Status    = "READY"
        }
    } catch {
        Write-Host "  [FAIL] $($t.Name): $_" -ForegroundColor Red
        $failed++
    }
}

foreach ($t in $fnfTasks) {
    try {
        $trigger = New-ScheduledTaskTrigger -Daily -At $t.TriggerDateTime
        Register-ScheduledTask `
            -TaskName $t.Name `
            -Action   $fnfAction `
            -Trigger  $trigger `
            -Settings $settings `
            -RunLevel Limited `
            -Force -ErrorAction Stop | Out-Null

        $registeredTasks += [PSCustomObject]@{
            Pipeline  = $t.Pipeline
            TaskName  = $t.Name
            Schedule  = "$($t.DisplayTime) Daily"
            CatchUp   = "Yes (On Boot)"
            Status    = "READY"
        }
    } catch {
        Write-Host "  [FAIL] $($t.Name): $_" -ForegroundColor Red
        $failed++
    }
}

foreach ($t in $elcmTasks) {
    try {
        $trigger = New-ScheduledTaskTrigger -Daily -At $t.TriggerDateTime
        Register-ScheduledTask `
            -TaskName $t.Name `
            -Action   $elcmAction `
            -Trigger  $trigger `
            -Settings $settings `
            -RunLevel Limited `
            -Force -ErrorAction Stop | Out-Null

        $registeredTasks += [PSCustomObject]@{
            Pipeline  = $t.Pipeline
            TaskName  = $t.Name
            Schedule  = "$($t.DisplayTime) Daily"
            CatchUp   = "Yes (On Boot)"
            Status    = "READY"
        }
    } catch {
        Write-Host "  [FAIL] $($t.Name): $_" -ForegroundColor Red
        $failed++
    }
}

Write-Host ""

# ==== Display Clean Tabular Output ====
if ($registeredTasks.Count -gt 0) {
    Write-Host "+-----------------+----------------------+------------------+-----------------+----------+" -ForegroundColor Green
    Write-Host "| PIPELINE        | TASK NAME            | SCHEDULE         | CATCH-UP (BOOT) | STATUS   |" -ForegroundColor Green
    Write-Host "+-----------------+----------------------+------------------+-----------------+----------+" -ForegroundColor Green
    foreach ($row in $registeredTasks) {
        $col1 = $row.Pipeline.PadRight(15)
        $col2 = $row.TaskName.PadRight(20)
        $col3 = $row.Schedule.PadRight(16)
        $col4 = $row.CatchUp.PadRight(15)
        $col5 = $row.Status.PadRight(8)
        Write-Host "| $col1 | $col2 | $col3 | $col4 | $col5 |" -ForegroundColor White
    }
    Write-Host "+-----------------+----------------------+------------------+-----------------+----------+" -ForegroundColor Green
}

if ($failed -eq 0) {
    Write-Host ""
    Write-Host "[OK] STATUS   : All $($registeredTasks.Count) tasks registered successfully." -ForegroundColor Green
    Write-Host "[OK] POLICY   : Tasks run silently in background via pythonw.exe." -ForegroundColor Green
    Write-Host "[OK] FAILOVER : Missed triggers (e.g. system shut down) fire automatically on next boot." -ForegroundColor Green
    Write-Host "[OK] RETRY    : If a run fails for any reason, Task Scheduler retries up to 3 times (5m interval)." -ForegroundColor Green
    Write-Host ""
    Write-Host "Quick Commands:" -ForegroundColor Gray
    Write-Host "  * Re-register tasks : powershell -ExecutionPolicy Bypass -File app\scheduler\setup_scheduler.ps1" -ForegroundColor Gray
    Write-Host "  * Uninstall tasks   : powershell -ExecutionPolicy Bypass -File app\scheduler\setup_scheduler.ps1 -Uninstall" -ForegroundColor Gray
    Write-Host "Log Locations:" -ForegroundColor Gray
    Write-Host "  * NDC Logs          : logs\download_ndc_report.log" -ForegroundColor Gray
    Write-Host "  * F&F Logs          : logs\process_fnf_closed_reports.log" -ForegroundColor Gray
    Write-Host "  * ELCM Logs         : logs\download_elcm_report.log" -ForegroundColor Gray
    Write-Host ""
} else {
    Write-Host "`n[WARNING] $failed task(s) failed to register." -ForegroundColor Yellow
}

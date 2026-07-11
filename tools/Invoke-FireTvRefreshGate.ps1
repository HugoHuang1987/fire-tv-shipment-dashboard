param(
    [string]$TriggerSource = "Windows任务计划",
    [string]$SharedRoot = $env:FIRE_TV_SHARED_ROOT
)

$ErrorActionPreference = "Stop"
if (-not $SharedRoot) {
    throw "FIRE_TV_SHARED_ROOT or -SharedRoot is required."
}
$Workspace = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$MonthlyScript = Join-Path $PSScriptRoot "Invoke-FireTvMonthlyRefresh.ps1"
$StatusScript = Join-Path $PSScriptRoot "Update-FireTvTaskStatus.ps1"
$LogsDir = Join-Path $SharedRoot "logs"
$LatestCheckPath = Join-Path $LogsDir "weekly-check-latest.json"
$CheckHistoryPath = Join-Path $LogsDir "weekly-check-history.jsonl"
$GateLockPath = Join-Path $LogsDir "weekly-check.lock"
$started = Get-Date
$monthStart = Get-Date -Year $started.Year -Month $started.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$monthKey = $monthStart.ToString("yyyyMM")
$successMarkerPath = Join-Path $LogsDir "monthly-success-$monthKey.json"
$finalExitCode = 1
$gateLock = $null

New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null

function Write-Utf8Text {
    param([string]$Path, [string]$Content)
    [System.IO.File]::WriteAllText($Path, $Content, (New-Object System.Text.UTF8Encoding($false)))
}

function Save-CheckRecord {
    param([System.Collections.IDictionary]$Record, [switch]$AppendHistory)
    $json = $Record | ConvertTo-Json -Depth 6 -Compress
    Write-Utf8Text -Path $LatestCheckPath -Content $json
    if ($AppendHistory) {
        [System.IO.File]::AppendAllText($CheckHistoryPath, "$json`r`n", (New-Object System.Text.UTF8Encoding($false)))
    }
}

function Get-ValidSuccessMarker {
    if (-not (Test-Path -LiteralPath $successMarkerPath)) {
        return $null
    }
    try {
        $marker = Get-Content -LiteralPath $successMarkerPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($marker.status -eq "success" -and [string]$marker.runMonth -eq $monthKey) {
            return $marker
        }
    } catch {}
    return $null
}

function Publish-TaskStatus {
    try {
        & $StatusScript -SharedRoot $SharedRoot | Out-Null
    } catch {
        $message = "[$(Get-Date -Format o)] 状态文件更新失败：$($_.Exception.Message)"
        Add-Content -LiteralPath (Join-Path $LogsDir "weekly-check-errors.log") -Value $message -Encoding UTF8
    }
}

$record = [ordered]@{
    schemaVersion = 1
    checkId = $started.ToString("yyyyMMdd_HHmmss")
    startedAt = $started.ToString("o")
    completedAt = ""
    source = $TriggerSource
    monthKey = $monthKey
    runDate = $monthStart.ToString("yyyy-MM-dd")
    decision = "checking"
    decisionLabel = "正在检查本月是否已经更新"
    result = "running"
    exitCode = $null
    durationSeconds = $null
    message = "Windows已启动本次轻量检查。"
    successMarker = $successMarkerPath
}

try {
    try {
        $gateLock = [System.IO.File]::Open($GateLockPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    } catch {
        throw "已有Fire TV每周检查正在运行。"
    }

    Save-CheckRecord -Record $record
    Publish-TaskStatus

    $marker = Get-ValidSuccessMarker
    if ($marker) {
        $record.decision = "skip_already_updated"
        $record.decisionLabel = "本月已经更新，未重复执行"
        $record.result = "success"
        $record.message = "已找到${monthKey}月成功标记，轻量检查完成。"
        $finalExitCode = 0
    } else {
        $record.decision = "refresh_required"
        $record.decisionLabel = "本月尚未更新，已启动完整刷新"
        $record.message = "未找到${monthKey}月成功标记，开始下载、合并、统计和发布。"
        Save-CheckRecord -Record $record
        Publish-TaskStatus

        & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $MonthlyScript -RunDate $monthStart.ToString("yyyy-MM-dd")
        $refreshExitCode = $LASTEXITCODE
        $record.exitCode = $refreshExitCode

        $markerAfterRefresh = Get-ValidSuccessMarker
        if ($refreshExitCode -eq 0 -and $markerAfterRefresh) {
            $record.decision = "refresh_completed"
            $record.decisionLabel = "本月尚未更新，已完成刷新"
            $record.result = "success"
            $record.message = "完整刷新成功，已生成${monthKey}月成功标记。"
            $finalExitCode = 0
        } elseif ($refreshExitCode -eq 2) {
            $record.decision = "refresh_blocked"
            $record.decisionLabel = "本月刷新被规则或数据冲突阻断"
            $record.result = "blocked"
            $record.message = "完整刷新被阻断，正式结果保留上一版。"
            $finalExitCode = 2
        } else {
            $record.decision = "refresh_failed"
            $record.decisionLabel = "本月刷新失败"
            $record.result = "failure"
            $record.message = if ($refreshExitCode -eq 0) { "刷新脚本返回成功，但未生成本月成功标记。" } else { "完整刷新失败，退出码：$refreshExitCode。" }
            $finalExitCode = 1
        }
    }
} catch {
    $record.decision = "gate_failed"
    $record.decisionLabel = "Windows轻量检查失败"
    $record.result = "failure"
    $record.message = $_.Exception.Message
    $finalExitCode = 1
} finally {
    $completed = Get-Date
    $record.completedAt = $completed.ToString("o")
    $record.durationSeconds = [int][Math]::Round(($completed - $started).TotalSeconds)
    if ($null -eq $record.exitCode) {
        $record.exitCode = $finalExitCode
    }
    try {
        Save-CheckRecord -Record $record -AppendHistory
    } catch {}
    Publish-TaskStatus
    if ($gateLock) {
        $gateLock.Dispose()
    }
}

exit $finalExitCode

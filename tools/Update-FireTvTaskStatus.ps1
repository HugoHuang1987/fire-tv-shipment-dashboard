param(
    [string]$SharedRoot = $env:FIRE_TV_SHARED_ROOT,
    [string]$DashboardSourceDir = "",
    [string]$TaskName = "Fire TV 出货量每周检查",
    [switch]$Show,
    [switch]$PassThru
)

$ErrorActionPreference = "Stop"
$Workspace = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
if (-not $SharedRoot) {
    throw "FIRE_TV_SHARED_ROOT or -SharedRoot is required."
}
if (-not $DashboardSourceDir) {
    $DashboardSourceDir = Join-Path $Workspace "dashboard"
}

$LogsDir = Join-Path $SharedRoot "logs"
$SharedDashboardDir = Join-Path $SharedRoot "dashboard"
$LatestCheckPath = Join-Path $LogsDir "weekly-check-latest.json"
$now = Get-Date

New-Item -ItemType Directory -Path $LogsDir, $DashboardSourceDir, $SharedDashboardDir -Force | Out-Null

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }
    try {
        return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        return $null
    }
}

function Convert-ToLocalDateTime {
    param($Value)
    if (-not $Value) {
        return $null
    }
    try {
        return ([datetimeoffset]::Parse([string]$Value)).LocalDateTime
    } catch {
        try {
            return [datetime]::Parse([string]$Value)
        } catch {
            return $null
        }
    }
}

function Format-DateTime {
    param($Value)
    if ($null -eq $Value) {
        return ""
    }
    try {
        $dateValue = [datetime]$Value
        if ($dateValue.Year -lt 2000) {
            return ""
        }
        return $dateValue.ToString("yyyy-MM-dd HH:mm:ss")
    } catch {
        return ""
    }
}

function Write-Utf8Text {
    param([string]$Path, [string]$Content)
    [System.IO.File]::WriteAllText($Path, $Content, (New-Object System.Text.UTF8Encoding($false)))
}

$task = $null
$taskInfo = $null
$taskError = ""
try {
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName -TaskPath $task.TaskPath -ErrorAction Stop
} catch {
    $taskError = $_.Exception.Message
}

$taskInstalled = $null -ne $task
$taskState = if ($taskInstalled) { [string]$task.State } else { "NotInstalled" }
$taskEnabled = $taskInstalled -and $taskState -ne "Disabled"
$lastTaskRunAt = if ($taskInfo) { Format-DateTime $taskInfo.LastRunTime } else { "" }
$nextTaskRunAt = if ($taskInfo) { Format-DateTime $taskInfo.NextRunTime } else { "" }
$lastTaskResultCode = if ($taskInfo) { [int64]$taskInfo.LastTaskResult } else { $null }
$lastTaskResultHex = ""
$lastTaskResultText = "尚未运行"
if ($null -ne $lastTaskResultCode) {
    $lastTaskResultHex = "0x{0:X8}" -f ([uint32]$lastTaskResultCode)
    switch ($lastTaskResultCode) {
        0 { $lastTaskResultText = "成功" }
        267009 { $lastTaskResultText = "正在运行" }
        267011 { $lastTaskResultText = "尚未运行" }
        default { $lastTaskResultText = "异常（$lastTaskResultHex）" }
    }
}

$registrationAt = $null
if ($taskInstalled -and $task.Date) {
    $registrationAt = Convert-ToLocalDateTime $task.Date
}

$latestCheck = Read-JsonFile $LatestCheckPath
$lastCheckedAtValue = $null
if ($latestCheck) {
    $lastCheckedAtValue = Convert-ToLocalDateTime $(if ($latestCheck.completedAt) { $latestCheck.completedAt } else { $latestCheck.startedAt })
}
$lastCheckedAt = Format-DateTime $lastCheckedAtValue

$monthStart = Get-Date -Year $now.Year -Month $now.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$monthKey = $monthStart.ToString("yyyyMM")
$currentMarkerPath = Join-Path $LogsDir "monthly-success-$monthKey.json"
$currentMarker = Read-JsonFile $currentMarkerPath
$currentMonthUpdated = $currentMarker -and $currentMarker.status -eq "success" -and [string]$currentMarker.runMonth -eq $monthKey

$latestSuccess = $null
$markerFiles = @(Get-ChildItem -LiteralPath $LogsDir -Filter "monthly-success-*.json" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
foreach ($markerFile in $markerFiles) {
    $candidate = Read-JsonFile $markerFile.FullName
    if ($candidate -and $candidate.status -eq "success") {
        $latestSuccess = $candidate
        break
    }
}

$daysSinceMonday = (([int]$now.DayOfWeek + 6) % 7)
$latestDue = $now.Date.AddDays(-$daysSinceMonday).AddHours(9)
if ($now -lt $latestDue) {
    $latestDue = $latestDue.AddDays(-7)
}

$health = "healthy"
$healthTitle = "计划任务运行正常"
$healthMessage = "每周一09:00检查；本月已更新时会直接跳过。"
if (-not $taskInstalled) {
    $health = "error"
    $healthTitle = "Windows计划任务未安装"
    $healthMessage = $taskError
} elseif (-not $taskEnabled) {
    $health = "error"
    $healthTitle = "Windows计划任务已停用"
    $healthMessage = "请重新启用任务后再检查。"
} else {
    $wasRegisteredBeforeDue = $null -eq $registrationAt -or $registrationAt -le $latestDue
    $missedLatestDue = $wasRegisteredBeforeDue -and ($null -eq $lastCheckedAtValue -or $lastCheckedAtValue -lt $latestDue)
    if ($missedLatestDue) {
        $health = "warning"
        $healthTitle = "最近一次计划检查尚未留下记录"
        $healthMessage = "计划时间为$($latestDue.ToString('yyyy-MM-dd HH:mm'))；Windows会在下次可用时补跑。"
    } elseif ($latestCheck -and $latestCheck.result -in @("failure", "blocked")) {
        $health = "warning"
        $healthTitle = "最近一次检查未完成正式发布"
        $healthMessage = [string]$latestCheck.message
    } elseif (-not $latestCheck) {
        $healthTitle = "计划任务已安装，等待首次检查"
        $healthMessage = "下次计划时间：$(if ($nextTaskRunAt) { $nextTaskRunAt } else { '等待Windows计算' })。"
    }
}

$payload = [ordered]@{
    schemaVersion = 1
    generatedAt = $now.ToString("yyyy-MM-dd HH:mm:ss")
    health = $health
    healthTitle = $healthTitle
    healthMessage = $healthMessage
    task = [ordered]@{
        name = $TaskName
        installed = $taskInstalled
        enabled = $taskEnabled
        state = $taskState
        schedule = "每周一 09:00"
        catchUp = $true
        lastRunAt = $lastTaskRunAt
        lastResultCode = $lastTaskResultHex
        lastResultText = $lastTaskResultText
        nextRunAt = $nextTaskRunAt
    }
    check = [ordered]@{
        lastCheckedAt = $lastCheckedAt
        source = if ($latestCheck) { [string]$latestCheck.source } else { "" }
        decision = if ($latestCheck) { [string]$latestCheck.decision } else { "not_run" }
        decisionLabel = if ($latestCheck) { [string]$latestCheck.decisionLabel } else { "尚未执行本机检查" }
        result = if ($latestCheck) { [string]$latestCheck.result } else { "not_run" }
        message = if ($latestCheck) { [string]$latestCheck.message } else { "" }
        durationSeconds = if ($latestCheck -and $null -ne $latestCheck.durationSeconds) { [int]$latestCheck.durationSeconds } else { $null }
    }
    month = [ordered]@{
        key = $monthKey
        updated = [bool]$currentMonthUpdated
        cutoffDate = if ($currentMonthUpdated) { [string]$currentMarker.cutoffDate } else { "" }
        completedAt = if ($currentMonthUpdated) { Format-DateTime (Convert-ToLocalDateTime $currentMarker.completedAt) } else { "" }
    }
    latestRefresh = [ordered]@{
        runMonth = if ($latestSuccess) { [string]$latestSuccess.runMonth } else { "" }
        cutoffDate = if ($latestSuccess) { [string]$latestSuccess.cutoffDate } else { "" }
        completedAt = if ($latestSuccess) { Format-DateTime (Convert-ToLocalDateTime $latestSuccess.completedAt) } else { "" }
    }
}

$json = $payload | ConvertTo-Json -Depth 7 -Compress
$sourceJson = Join-Path $DashboardSourceDir "scheduler-status.json"
$sourceJs = Join-Path $DashboardSourceDir "scheduler-status.js"
$sharedJson = Join-Path $SharedDashboardDir "scheduler-status.json"
$sharedJs = Join-Path $SharedDashboardDir "scheduler-status.js"

Write-Utf8Text -Path $sourceJson -Content $json
Write-Utf8Text -Path $sourceJs -Content "window.FTV_SCHEDULER_STATUS=$json;`n"
Copy-Item -LiteralPath $sourceJson -Destination $sharedJson -Force
Copy-Item -LiteralPath $sourceJs -Destination $sharedJs -Force

if ($Show) {
    Write-Host ""
    Write-Host "Fire TV 出货量任务状态"
    Write-Host "任务状态：$healthTitle"
    Write-Host "是否启用：$(if ($taskEnabled) { '是' } else { '否' })"
    Write-Host "最近检测：$(if ($lastCheckedAt) { $lastCheckedAt } else { '尚无记录' })"
    Write-Host "检测结果：$($payload.check.decisionLabel)"
    Write-Host "最近刷新：$(if ($payload.latestRefresh.completedAt) { $payload.latestRefresh.completedAt } else { '尚无记录' })"
    Write-Host "上次任务：$(if ($lastTaskRunAt) { "$lastTaskRunAt / $lastTaskResultText" } else { '尚未运行' })"
    Write-Host "下次任务：$(if ($nextTaskRunAt) { $nextTaskRunAt } else { '尚未计算' })"
    Write-Host "错过补跑：已启用"
    if ($healthMessage) {
        Write-Host "说明：$healthMessage"
    }
    Write-Host ""
}

if ($PassThru) {
    return [pscustomobject]$payload
}

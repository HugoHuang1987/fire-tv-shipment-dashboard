param(
    [datetime]$RunDate = (Get-Date -Day 1).Date,
    [string]$SharedRoot = $env:FIRE_TV_SHARED_ROOT,
    [string]$Python = $env:FIRE_TV_PYTHON,
    [string]$ExistingDownloadDir = "",
    [switch]$SkipNotification,
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$Workspace = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$DashboardDir = Join-Path $Workspace "dashboard"
if (-not $SharedRoot) { throw "FIRE_TV_SHARED_ROOT or -SharedRoot is required." }
if (-not $Python) { $Python = "python" }
$NotifyScript = Join-Path $DashboardDir "Send-FireTvNotification.ps1"
$PersonnelFile = Join-Path $SharedRoot "人员名单.txt"
$EngineerIdFile = Join-Path $SharedRoot "Fire TV quantity - engineer-ids.csv"
$RawPath = Join-Path $SharedRoot "Fire TV quantity - raw.xlsx"
$LogsDir = Join-Path $SharedRoot "logs"
$HistoryRawDir = Join-Path $SharedRoot "history\raw"
$SnapshotRoot = Join-Path $HistoryRawDir "rdm-snapshots"

$RunDate = $RunDate.Date
$monthKey = $RunDate.ToString("yyyyMM")
$cutoffText = $RunDate.ToString("yyyy-MM-dd")
$startDate = $RunDate.AddMonths(-8).ToString("yyyy/MM/dd")
$endDate = $RunDate.ToString("yyyy/MM/dd")
$successMarker = Join-Path $LogsDir "monthly-success-$monthKey.json"
$lockPath = Join-Path $LogsDir "monthly-refresh.lock"
$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$runLog = Join-Path $LogsDir "monthly-refresh-$runStamp.log"
$mergeReport = Join-Path $LogsDir "source-merge-$($RunDate.ToString('yyyyMMdd'))-$runStamp.json"
$conflictWorkbook = Join-Path $SharedRoot "Fire TV quantity - source-conflicts.xlsx"

New-Item -ItemType Directory -Path $LogsDir, $HistoryRawDir, $SnapshotRoot -Force | Out-Null

function Send-RefreshNotification {
    param(
        [ValidateSet("Success", "Warning", "Blocked", "Failure")]
        [string]$Level,
        [string]$Title,
        [string]$Message,
        [string]$OpenPath = ""
    )
    if ($SkipNotification) {
        return
    }
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $NotifyScript -Level $Level -Title $Title -Message $Message -OpenPath $OpenPath | Out-String | Add-Content -LiteralPath $runLog -Encoding UTF8
    } catch {
        "Notification failed: $($_.Exception.Message)" | Add-Content -LiteralPath $runLog -Encoding UTF8
    }
}

function Set-DashboardRunStatus {
    param(
        [string]$Level,
        [string]$Title,
        [string]$Message,
        [string]$RelatedFile = "",
        [object[]]$Warnings = @()
    )
    $attemptedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $cleanWarnings = @($Warnings | Where-Object { $null -ne $_ })
    $payload = [ordered]@{
        level = $Level
        title = $Title
        message = $Message
        cutoffDate = $cutoffText
        attemptedAt = $attemptedAt
        statusKey = "$Level|$cutoffText|$attemptedAt"
        relatedFile = $RelatedFile
        warnings = $cleanWarnings
    }
    $json = $payload | ConvertTo-Json -Depth 5 -Compress
    $localJson = Join-Path $DashboardDir "dashboard-run-status.json"
    $localJs = Join-Path $DashboardDir "dashboard-run-status.js"
    [System.IO.File]::WriteAllText($localJson, $json, (New-Object System.Text.UTF8Encoding($true)))
    [System.IO.File]::WriteAllText($localJs, "window.FTV_RUN_STATUS=$json;`n", (New-Object System.Text.UTF8Encoding($true)))
    $sharedDashboard = Join-Path $SharedRoot "dashboard"
    New-Item -ItemType Directory -Path $sharedDashboard -Force | Out-Null
    Copy-Item -LiteralPath $localJson -Destination (Join-Path $sharedDashboard "dashboard-run-status.json") -Force
    Copy-Item -LiteralPath $localJs -Destination (Join-Path $sharedDashboard "dashboard-run-status.js") -Force
    Copy-Item -LiteralPath (Join-Path $DashboardDir "index.html") -Destination (Join-Path $sharedDashboard "index.html") -Force
    New-Item -ItemType Directory -Path (Join-Path $sharedDashboard "assets") -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $DashboardDir "assets\styles.css") -Destination (Join-Path $sharedDashboard "assets\styles.css") -Force
    Copy-Item -LiteralPath (Join-Path $DashboardDir "assets\app.js") -Destination (Join-Path $sharedDashboard "assets\app.js") -Force
    Copy-Item -LiteralPath (Join-Path $DashboardDir "assets\lucide.min.js") -Destination (Join-Path $sharedDashboard "assets\lucide.min.js") -Force
}

function Get-MissingEngineerWarning {
    if (-not (Test-Path -LiteralPath $PersonnelFile) -or -not (Test-Path -LiteralPath $EngineerIdFile)) {
        return $null
    }
    $personnel = @(Get-Content -LiteralPath $PersonnelFile -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $idRows = @(Import-Csv -LiteralPath $EngineerIdFile -Encoding UTF8)
    $known = @{}
    foreach ($row in $idRows) {
        $name = ([string]$row.name).Trim()
        $id = ([string]$row.engineer_id).Trim()
        if ($name -and $id) {
            $known[$name] = $true
        }
    }
    $missing = @($personnel | Where-Object { -not $known.ContainsKey($_) })
    if ($missing.Count -eq 0) {
        return $null
    }
    $personnelHash = (Get-FileHash -LiteralPath $PersonnelFile -Algorithm SHA256).Hash
    return [ordered]@{
        id = "missing_engineer_ids"
        level = "warning"
        title = "人员ID未匹配"
        message = "$($missing -join '、')仍无RDM历史和ID，已按确认规则继续，不阻断发布。"
        people = $missing
        version = $personnelHash
        resetWhen = "人员名单.txt内容变化"
    }
}

function Copy-DirectoryWithRetry {
    param(
        [string]$SourceDir,
        [string]$DestinationDir,
        [int]$Attempts = 12,
        [int]$DelaySec = 5
    )

    New-Item -ItemType Directory -Path $DestinationDir -Force | Out-Null
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            Get-ChildItem -LiteralPath $SourceDir -Force | ForEach-Object {
                Copy-Item -LiteralPath $_.FullName -Destination $DestinationDir -Recurse -Force
            }
            return
        } catch {
            if ($attempt -eq $Attempts) {
                throw
            }
            Start-Sleep -Seconds $DelaySec
        }
    }
}

if ((Test-Path -LiteralPath $successMarker) -and -not $Force) {
    Write-Host "本月已成功刷新：$successMarker"
    exit 0
}

$lockStream = $null
try {
    try {
        $lockStream = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $lockBytes = [System.Text.Encoding]::UTF8.GetBytes("$PID $(Get-Date -Format o)")
        $lockStream.Write($lockBytes, 0, $lockBytes.Length)
        $lockStream.Flush()
    } catch {
        throw "已有出货量刷新任务正在运行：$lockPath"
    }

    "[$(Get-Date -Format o)] START cutoff=$cutoffText window=$startDate..$endDate" | Set-Content -LiteralPath $runLog -Encoding UTF8

    if ($ExistingDownloadDir) {
        $downloadDir = (Resolve-Path -LiteralPath $ExistingDownloadDir).Path
    } else {
        $downloadDir = Join-Path $SnapshotRoot "rdm_$($RunDate.ToString('yyyyMMdd'))_$runStamp"
        $downloadWorkDir = Join-Path ([System.IO.Path]::GetTempPath()) "fire_tv_rdm_$($RunDate.ToString('yyyyMMdd'))_$runStamp"
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "Invoke-RdmDownload.ps1") `
            -EngineersFile $PersonnelFile -EngineerIdsCsv $EngineerIdFile -OnMissing Skip `
            -StartDate $startDate -EndDate $endDate -PageSize 200 -OutputDir $downloadWorkDir *>&1 | Tee-Object -FilePath $runLog -Append
        $downloadExit = $LASTEXITCODE
        Copy-DirectoryWithRetry -SourceDir $downloadWorkDir -DestinationDir $downloadDir
        [System.IO.File]::WriteAllText((Join-Path $LogsDir "latest-rdm-download-path.txt"), $downloadDir, (New-Object System.Text.UTF8Encoding($true)))
        "Published RDM snapshot: $downloadDir" | Add-Content -LiteralPath $runLog -Encoding UTF8
        try {
            Remove-Item -LiteralPath $downloadWorkDir -Recurse -Force
        } catch {}
        if ($downloadExit -ne 0) {
            throw "RDM下载失败，退出码：$downloadExit"
        }
    }

    $snapshotPath = Join-Path $downloadDir "data.csv"
    $summaryPath = Join-Path $downloadDir "summary.csv"
    if (-not (Test-Path -LiteralPath $snapshotPath) -or -not (Test-Path -LiteralPath $summaryPath)) {
        throw "RDM下载缺少data.csv或summary.csv：$downloadDir"
    }
    $failedPeople = @(Import-Csv -LiteralPath $summaryPath -Encoding UTF8 | Where-Object { $_.status -ne "OK" })
    if ($failedPeople.Count -gt 0) {
        throw "RDM人员下载存在失败：$($failedPeople.engineer -join ', ')"
    }

    $mergeScript = Join-Path $PSScriptRoot "merge_rdm_snapshot.py"
    $mergeArguments = @(
        "`"$mergeScript`"",
        "--raw", "`"$RawPath`"",
        "--snapshot", "`"$snapshotPath`"",
        "--report", "`"$mergeReport`"",
        "--backup-dir", "`"$HistoryRawDir`"",
        "--cutoff", $cutoffText,
        "--conflict-xlsx", "`"$conflictWorkbook`""
    )
    $mergeProcess = Start-Process -FilePath $Python -ArgumentList $mergeArguments -Wait -PassThru -WindowStyle Hidden
    $mergeExit = $mergeProcess.ExitCode
    "Merge process exit code: $mergeExit" | Add-Content -LiteralPath $runLog -Encoding UTF8
    $merge = Get-Content -LiteralPath $mergeReport -Raw -Encoding UTF8 | ConvertFrom-Json
    $missingEngineerWarning = Get-MissingEngineerWarning

    if ($mergeExit -eq 2 -or $merge.status -eq "blocked") {
        $message = "$($RunDate.ToString('yyyy年MM月'))补跑已下载$($merge.downloadedUniqueOrders)张唯一订单，但发现$($merge.conflictCount)条锁定字段冲突，正式raw、统计结果和页面均保留上一版。"
        Set-DashboardRunStatus -Level "blocked" -Title "本月刷新被阻断，当前仍显示上一版正式数据" -Message $message -RelatedFile $conflictWorkbook -Warnings @($missingEngineerWarning)
        Send-RefreshNotification -Level "Blocked" -Title "Fire TV出货量月度刷新被阻断" -Message $message -OpenPath $conflictWorkbook
        exit 2
    }
    if ($mergeExit -ne 0) {
        throw "原始数据滚动合并失败，退出码：$mergeExit"
    }

    $env:FIRE_TV_NO_PAUSE = "1"
    $env:FIRE_TV_DISABLE_PATCH = "1"
    $env:FIRE_TV_DATA_AS_OF = $cutoffText
    $env:FIRE_TV_SKIP_WIKI = "1"
    & (Join-Path $DashboardDir "刷新看板数据.cmd") $SharedRoot *>&1 | Tee-Object -FilePath $runLog -Append
    $refreshExit = $LASTEXITCODE
    if ($refreshExit -ne 0) {
        throw "统计或页面刷新失败，退出码：$refreshExit"
    }

    $dashboardData = Get-Content -LiteralPath (Join-Path $DashboardDir "dashboard-data.json") -Raw -Encoding UTF8 | ConvertFrom-Json
    $summary = $dashboardData.audit.summary
    $markerPayload = [ordered]@{
        schemaVersion = 1
        status = "success"
        runMonth = $monthKey
        cutoffDate = $cutoffText
        completedAt = (Get-Date).ToString("o")
        downloadDir = $downloadDir
        downloadedUniqueOrders = $merge.downloadedUniqueOrders
        updatedOrders = $merge.updatedOrders
        addedOrders = $merge.addedOrders
        includedQuantity = $summary.includedQuantity
        ruleVersion = $dashboardData.calculationRuleVersion
        dashboard = (Join-Path $SharedRoot "dashboard\index.html")
    }
    $markerPayload | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $successMarker -Encoding UTF8
    $message = "截止$cutoffText刷新成功：RDM $($merge.downloadedUniqueOrders)张唯一订单，新增$($merge.addedOrders)张、更新$($merge.updatedOrders)张；计入出货量$($summary.includedQuantity)台；规则$($dashboardData.calculationRuleVersion)。"
    Set-DashboardRunStatus -Level "success" -Title "本月刷新成功" -Message $message -Warnings @($missingEngineerWarning)
    Send-RefreshNotification -Level "Success" -Title "Fire TV出货量月度刷新成功" -Message $message -OpenPath (Join-Path $SharedRoot "dashboard\index.html")
    exit 0
} catch {
    $errorMessage = $_.Exception.Message
    "[$(Get-Date -Format o)] FAILURE $errorMessage" | Add-Content -LiteralPath $runLog -Encoding UTF8
    Set-DashboardRunStatus -Level "failure" -Title "本月刷新失败，当前仍显示上一版正式数据" -Message $errorMessage -RelatedFile $runLog
    Send-RefreshNotification -Level "Failure" -Title "Fire TV出货量月度刷新失败" -Message "$errorMessage。正式结果保持上一版。" -OpenPath $runLog
    exit 1
} finally {
    if ($lockStream) {
        $lockStream.Dispose()
    }
    if (Test-Path -LiteralPath $lockPath) {
        Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
    }
}

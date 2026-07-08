param(
    [string]$WikiUrl = $env:FIRE_TV_WIKI_URL,
    [string]$RulePath = $env:FIRE_TV_RULE_PATH,
    [string]$CalculationRulePath = '',
    [string]$CachedHtmlPath = '',
    [string]$RunRoot = ''
)

$ErrorActionPreference = 'Stop'
if (-not $WikiUrl) { throw 'FIRE_TV_WIKI_URL or -WikiUrl is required.' }
if (-not $RulePath) { throw 'FIRE_TV_RULE_PATH or -RulePath is required.' }
$Python = if ($env:FIRE_TV_PYTHON) { $env:FIRE_TV_PYTHON } else { 'python' }
$Node = if ($env:FIRE_TV_NODE) { $env:FIRE_TV_NODE } else { 'node' }
$SharedRoot = Split-Path -Parent $RulePath
if (-not $CalculationRulePath) { $CalculationRulePath = Join-Path $SharedRoot 'Fire TV quantity - calculation-rules.json' }
if (-not $RunRoot) { $RunRoot = Join-Path $PSScriptRoot 'sync_runs' }
$Stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$RunDir = Join-Path $RunRoot $Stamp
$HtmlPath = Join-Path $RunDir 'wiki.html'
$PayloadPath = Join-Path $RunDir 'wiki_rules.json'
$OutputPath = Join-Path $RunDir 'Fire TV quantity - rule.updated.xlsx'
$ReportPath = Join-Path $RunDir 'sync_report.json'
$AlertTempPath = Join-Path $RunDir 'Fire TV quantity - wiki-sync-alerts.xlsx'
$QaDir = Join-Path $RunDir 'qa'
$AlertPath = Join-Path $SharedRoot 'Fire TV quantity - wiki-sync-alerts.xlsx'
$HistoryDir = Join-Path $SharedRoot 'history\rules'
$ReportDir = Join-Path $SharedRoot 'logs\wiki-sync'
$SavedReport = Join-Path $ReportDir ("wiki_sync_{0}.json" -f $Stamp)
$SavedAlert = Join-Path $ReportDir ("wiki_sync_{0}.xlsx" -f $Stamp)
$LatestStatus = Join-Path $ReportDir 'latest_status.json'

function Write-SyntheticFailureReport {
    param([string]$Type, [string]$Message)
    $Synthetic = [ordered]@{
        schemaVersion = 2
        status = 'technical_error'
        sourceUrl = $WikiUrl
        extractedAt = ''
        generatedAt = (Get-Date).ToString('o')
        workbook = $RulePath
        worksheet = ''
        ruleUpdated = $false
        ruleProjectCount = 0
        ruleEngineCount = 0
        wikiRows = 0
        matchedWikiRecords = 0
        safeAdditionCount = 0
        blockingCount = 1
        warningCount = 0
        legacyProjectCount = 0
        blocking = @([ordered]@{
            type = $Type
            severity = '阻断'
            message = $Message
            resolution = '根据错误提示恢复网络、登录状态或文件访问后重新刷新。'
        })
        warnings = @()
        safeAdditions = @()
        changes = @()
        legacyCodesRetained = @()
        wikiRowsSnapshot = @()
    }
    $Json = $Synthetic | ConvertTo-Json -Depth 20
    [IO.File]::WriteAllText($ReportPath, $Json, (New-Object System.Text.UTF8Encoding($false)))
}

function Publish-AuditArtifacts {
    $SuccessMarker = "$AlertTempPath.success"
    Remove-Item -LiteralPath $SuccessMarker -Force -ErrorAction SilentlyContinue
    & $Node (Join-Path $PSScriptRoot 'build_wiki_sync_alerts.mjs') $ReportPath $AlertTempPath $QaDir
    if (-not (Test-Path -LiteralPath $SuccessMarker) -or -not (Test-Path -LiteralPath $AlertTempPath)) {
        throw 'Wiki同步告警表生成失败。'
    }
    Copy-Item -LiteralPath $ReportPath -Destination $SavedReport -Force
    Copy-Item -LiteralPath $ReportPath -Destination $LatestStatus -Force
    Copy-Item -LiteralPath $AlertTempPath -Destination $SavedAlert -Force
    Copy-Item -LiteralPath $AlertTempPath -Destination $AlertPath -Force
}

New-Item -ItemType Directory -Force -Path $RunDir,$HistoryDir,$ReportDir | Out-Null

try {
    if (-not (Test-Path -LiteralPath $RulePath)) { throw "规则表不存在：$RulePath" }
    if (-not (Test-Path -LiteralPath $CalculationRulePath)) { throw "统计规则不存在：$CalculationRulePath" }
    if (-not (Test-Path -LiteralPath $Python)) { throw "未找到Python运行环境：$Python" }
    if (-not (Test-Path -LiteralPath $Node)) { throw "未找到Node运行环境：$Node" }

    if ($CachedHtmlPath) {
        if (-not (Test-Path -LiteralPath $CachedHtmlPath)) { throw "缓存HTML不存在：$CachedHtmlPath" }
        Copy-Item -LiteralPath $CachedHtmlPath -Destination $HtmlPath -Force
    } else {
        $Client = $null
        $Handler = $null
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Add-Type -AssemblyName System.Net.Http
            $Handler = New-Object System.Net.Http.HttpClientHandler
            $Handler.UseProxy = $false
            $Handler.UseDefaultCredentials = $true
            $Client = New-Object System.Net.Http.HttpClient($Handler)
            $Client.Timeout = [TimeSpan]::FromSeconds(60)
            $Response = $Client.GetAsync($WikiUrl).GetAwaiter().GetResult()
            $Response.EnsureSuccessStatusCode()
            $Bytes = $Response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
            [IO.File]::WriteAllBytes($HtmlPath, $Bytes)
        } catch {
            throw "无法访问Wiki页面。请确认公司网络、登录状态或VPN后重试。原始错误：$($_.Exception.Message)"
        } finally {
            if ($null -ne $Client) { $Client.Dispose() }
            if ($null -ne $Handler) { $Handler.Dispose() }
        }
    }

    & $Python (Join-Path $PSScriptRoot 'extract_wiki_rule_table.py') --html $HtmlPath --output $PayloadPath --url $WikiUrl
    if ($LASTEXITCODE -ne 0) { throw 'Wiki表格解析失败，规则表未修改。' }

    & $Node (Join-Path $PSScriptRoot 'sync_rule_workbook.mjs') $PayloadPath $RulePath $OutputPath $ReportPath $CalculationRulePath
    if ($LASTEXITCODE -ne 0) { throw '规则工作簿审计失败，原文件未修改。' }

    Publish-AuditArtifacts
    $Report = Get-Content -LiteralPath $ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($Report.blockingCount -gt 0) {
        Write-Host ''
        Write-Host "Wiki同步发现 $($Report.blockingCount) 条阻断异常，正式规则未修改。" -ForegroundColor Red
        Write-Host "请查看：$AlertPath" -ForegroundColor Yellow
        exit 2
    }

    $BackupPath = ''
    if ($Report.ruleUpdated) {
        $BackupPath = Join-Path $HistoryDir ("Fire TV quantity - rule_{0}.xlsx" -f $Stamp)
        $StagedRulePath = "$RulePath.sync-new"
        Copy-Item -LiteralPath $RulePath -Destination $BackupPath -Force
        Copy-Item -LiteralPath $OutputPath -Destination $StagedRulePath -Force
        Move-Item -LiteralPath $StagedRulePath -Destination $RulePath -Force
    }

    Write-Host "Wiki同步通过。匹配记录：$($Report.matchedWikiRecords)，安全新增：$($Report.safeAdditionCount)，提示：$($Report.warningCount)。"
    if ($BackupPath) { Write-Host "原规则备份：$BackupPath" }
    Write-Host "同步审计：$AlertPath"
    Write-Host "详细记录：$SavedReport"
    exit 0
} catch {
    $FailureMessage = $_.Exception.Message
    try {
        if (-not (Test-Path -LiteralPath $ReportPath)) {
            Write-SyntheticFailureReport -Type 'wiki_sync_technical_error' -Message $FailureMessage
        }
        Publish-AuditArtifacts
    } catch {
        Write-Host "告警文件也未能生成：$($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Host ''
    Write-Host "Wiki同步失败，正式规则未修改：$FailureMessage" -ForegroundColor Red
    if (Test-Path -LiteralPath $AlertPath) { Write-Host "请查看：$AlertPath" -ForegroundColor Yellow }
    exit 1
}

param(
    [ValidateSet('Test', 'Success', 'Warning', 'Blocked', 'Failure')]
    [string]$Level = 'Test',
    [string]$Title = 'Fire TV 出货量',
    [string]$Message = '通知通道测试成功',
    [string]$OpenPath = '',
    [switch]$SkipWindows,
    [switch]$SkipServerChan,
    [switch]$SkipWeCom
)

$ErrorActionPreference = 'Stop'
$results = @()

function Add-Result {
    param([string]$Channel, [string]$Status, [string]$Detail)
    $script:results += [pscustomobject]@{
        Channel = $Channel
        Status = $Status
        Detail = $Detail
    }
}

function Send-WindowsBalloon {
    param([string]$NotificationTitle, [string]$NotificationMessage, [string]$NotificationLevel)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $notify = New-Object System.Windows.Forms.NotifyIcon
    try {
        $notify.Icon = switch ($NotificationLevel) {
            'Success' { [System.Drawing.SystemIcons]::Information }
            'Warning' { [System.Drawing.SystemIcons]::Warning }
            'Blocked' { [System.Drawing.SystemIcons]::Warning }
            'Failure' { [System.Drawing.SystemIcons]::Error }
            default { [System.Drawing.SystemIcons]::Information }
        }
        $notify.Visible = $true
        $notify.BalloonTipTitle = $NotificationTitle
        $notify.BalloonTipText = $NotificationMessage
        $notify.BalloonTipIcon = switch ($NotificationLevel) {
            'Success' { [System.Windows.Forms.ToolTipIcon]::Info }
            'Warning' { [System.Windows.Forms.ToolTipIcon]::Warning }
            'Blocked' { [System.Windows.Forms.ToolTipIcon]::Warning }
            'Failure' { [System.Windows.Forms.ToolTipIcon]::Error }
            default { [System.Windows.Forms.ToolTipIcon]::Info }
        }
        $notify.ShowBalloonTip(8000)
        Start-Sleep -Seconds 3
        Add-Result -Channel 'Windows' -Status 'Sent' -Detail '已发送到当前Windows登录会话'
    } finally {
        $notify.Dispose()
    }
}

function Send-WeComRobot {
    param([string]$NotificationTitle, [string]$NotificationMessage, [string]$NotificationLevel, [string]$RelatedPath)
    $webhook = [Environment]::GetEnvironmentVariable('FIRE_TV_WECOM_WEBHOOK', 'Process')
    if (-not $webhook) { $webhook = [Environment]::GetEnvironmentVariable('FIRE_TV_WECOM_WEBHOOK', 'User') }
    if (-not $webhook) {
        Add-Result -Channel '企业微信群机器人' -Status 'Skipped' -Detail '尚未配置FIRE_TV_WECOM_WEBHOOK'
        return
    }
    $levelLabel = @{
        Test = '测试'
        Success = '成功'
        Warning = '提示'
        Blocked = '阻断'
        Failure = '失败'
    }[$NotificationLevel]
    $lines = @(
        "### $NotificationTitle",
        "> 状态：$levelLabel",
        "> 时间：$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))",
        "> $NotificationMessage"
    )
    if ($RelatedPath) { $lines += "> 本地文件：$RelatedPath" }
    $payload = @{
        msgtype = 'markdown'
        markdown = @{ content = ($lines -join "`n") }
    } | ConvertTo-Json -Depth 5
    $response = Invoke-RestMethod -Method Post -Uri $webhook -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($payload)) -TimeoutSec 20
    if ($null -ne $response.errcode -and [int]$response.errcode -ne 0) {
        throw "企业微信返回错误：$($response.errcode) $($response.errmsg)"
    }
    Add-Result -Channel '企业微信群机器人' -Status 'Sent' -Detail '企业微信群消息已发送'
}

function Send-ServerChan {
    param([string]$NotificationTitle, [string]$NotificationMessage, [string]$NotificationLevel, [string]$RelatedPath)
    $sendKey = [Environment]::GetEnvironmentVariable('FIRE_TV_SERVERCHAN_SENDKEY', 'Process')
    if (-not $sendKey) { $sendKey = [Environment]::GetEnvironmentVariable('FIRE_TV_SERVERCHAN_SENDKEY', 'User') }
    if (-not $sendKey) {
        Add-Result -Channel 'Server酱个人微信' -Status 'Skipped' -Detail '尚未配置FIRE_TV_SERVERCHAN_SENDKEY'
        return
    }
    $levelLabel = @{
        Test = '测试'
        Success = '成功'
        Warning = '提示'
        Blocked = '阻断'
        Failure = '失败'
    }[$NotificationLevel]
    $details = @(
        "状态：$levelLabel",
        "时间：$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))",
        '',
        $NotificationMessage
    )
    if ($RelatedPath) { $details += @('', "本地文件：$RelatedPath") }
    $uri = "https://sctapi.ftqq.com/$sendKey.send"
    $response = Invoke-RestMethod -Method Post -Uri $uri -Body @{
        title = "[$levelLabel] $NotificationTitle"
        desp = ($details -join "`n")
    } -ContentType 'application/x-www-form-urlencoded; charset=utf-8' -TimeoutSec 20
    if ($null -ne $response.code -and [int]$response.code -ne 0) {
        throw "Server酱返回错误：$($response.code) $($response.message)"
    }
    Add-Result -Channel 'Server酱个人微信' -Status 'Sent' -Detail '个人微信消息已提交'
}

if (-not $SkipWindows) {
    try {
        Send-WindowsBalloon -NotificationTitle $Title -NotificationMessage $Message -NotificationLevel $Level
    } catch {
        Add-Result -Channel 'Windows' -Status 'Failed' -Detail $_.Exception.Message
    }
}

if (-not $SkipServerChan) {
    try {
        Send-ServerChan -NotificationTitle $Title -NotificationMessage $Message -NotificationLevel $Level -RelatedPath $OpenPath
    } catch {
        Add-Result -Channel 'Server酱个人微信' -Status 'Failed' -Detail $_.Exception.Message
    }
}

if (-not $SkipWeCom) {
    try {
        Send-WeComRobot -NotificationTitle $Title -NotificationMessage $Message -NotificationLevel $Level -RelatedPath $OpenPath
    } catch {
        Add-Result -Channel '企业微信群机器人' -Status 'Failed' -Detail $_.Exception.Message
    }
}

$results
if ($results.Status -contains 'Failed') { exit 2 }
exit 0

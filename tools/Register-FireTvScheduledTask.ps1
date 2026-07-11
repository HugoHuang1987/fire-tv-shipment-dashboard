param(
    [string]$TaskName = "Fire TV 出货量每周检查",
    [string]$SharedRoot = $env:FIRE_TV_SHARED_ROOT
)

$ErrorActionPreference = "Stop"
if (-not $SharedRoot) {
    throw "FIRE_TV_SHARED_ROOT or -SharedRoot is required."
}
$GateScript = Join-Path $PSScriptRoot "Invoke-FireTvRefreshGate.ps1"
$StatusScript = Join-Path $PSScriptRoot "Update-FireTvTaskStatus.ps1"
$PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"

if (-not (Test-Path -LiteralPath $GateScript)) {
    throw "缺少轻量检查脚本：$GateScript"
}

$actionArguments = "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$GateScript`" -TriggerSource `"Windows任务计划`" -SharedRoot `"$SharedRoot`""
$action = New-ScheduledTaskAction -Execute $PowerShellExe -Argument $actionArguments
$trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek Monday -At 9:00AM
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -WakeToRun `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 4)

$userId = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited
$description = "每周一09:00检查Fire TV出货量本月是否已更新；未更新才执行完整刷新。错过时间后在Windows下次可用时补跑。"
$task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description $description
Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null

& $StatusScript -SharedRoot $SharedRoot | Out-Null

$registered = Get-ScheduledTask -TaskName $TaskName
$info = Get-ScheduledTaskInfo -TaskName $TaskName -TaskPath $registered.TaskPath
Write-Host "Windows计划任务安装完成。"
Write-Host "任务名称：$TaskName"
Write-Host "计划时间：每周一 09:00"
Write-Host "错过补跑：已启用"
Write-Host "下次运行：$($info.NextRunTime.ToString('yyyy-MM-dd HH:mm:ss'))"

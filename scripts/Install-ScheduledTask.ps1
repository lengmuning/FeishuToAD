# Install-ScheduledTask.ps1 —— 注册每小时全量同步的任务计划
#
# 用法（以管理员身份在服务器上跑）：
#   .\scripts\Install-ScheduledTask.ps1
#   .\scripts\Install-ScheduledTask.ps1 -User "DOMAIN\ServiceAccount" -Password "xxxxx"
#   .\scripts\Install-ScheduledTask.ps1 -Remove    # 卸载任务
#
# 任务：FeishuToAD-Sync    每小时整点执行 Sync-FeishuToAD.ps1 -Mode Full

[CmdletBinding()]
param(
    [string]$TaskName = 'FeishuToAD-Sync',
    [string]$User,
    [string]$Password,
    [switch]$Remove
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
$mainScript = Join-Path $root 'Sync-FeishuToAD.ps1'

if ($Remove) {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "[OK] 任务已卸载: $TaskName" -ForegroundColor Green
    } else {
        Write-Host "[INFO] 任务不存在: $TaskName" -ForegroundColor Gray
    }
    return
}

if (-not (Test-Path $mainScript)) {
    throw "主脚本不存在: $mainScript"
}

# 要求以管理员跑
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    throw "此脚本必须以管理员身份运行（右键 PowerShell → 以管理员身份运行）"
}

# 询问运行账号
if (-not $User) {
    Write-Host ""
    Write-Host "任务需要以域管理员身份运行（可以创建 AD 用户/OU）。" -ForegroundColor Yellow
    $User = Read-Host "请输入运行账号（格式 域\账号，例如 DOMAIN\ServiceAccount）"
}
if (-not $Password) {
    $sec = Read-Host "请输入 $User 的密码" -AsSecureString
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
    $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
}

# 已存在则先卸载
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Write-Host "[INFO] 已存在同名任务，先卸载..." -ForegroundColor Gray
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# 构造任务
$action = New-ScheduledTaskAction `
    -Execute 'PowerShell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$mainScript`" -Mode Full" `
    -WorkingDirectory $root

# 每小时整点：Once + 重复间隔 1 小时，持续 10 年
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).Date `
    -RepetitionInterval (New-TimeSpan -Hours 1) `
    -RepetitionDuration (New-TimeSpan -Days 3650)

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

Register-ScheduledTask -TaskName $TaskName `
    -Action $action -Trigger $trigger -Settings $settings `
    -User $User -Password $Password `
    -RunLevel Highest `
    -Description 'Feishu 组织架构同步到 AD 域控。每小时整点执行 Sync-FeishuToAD.ps1 -Mode Full'

Write-Host ""
Write-Host "[OK] 任务已注册: $TaskName" -ForegroundColor Green
Write-Host "    运行账号: $User"
Write-Host "    脚本:    $mainScript -Mode Full"
Write-Host "    触发:    每小时整点"
Write-Host ""
Write-Host "手动立刻触发一次：" -ForegroundColor Cyan
Write-Host "    Start-ScheduledTask -TaskName '$TaskName'"
Write-Host ""
Write-Host "查看任务状态：" -ForegroundColor Cyan
Write-Host "    Get-ScheduledTaskInfo -TaskName '$TaskName'"
Write-Host ""
Write-Host "卸载：" -ForegroundColor Cyan
Write-Host "    .\scripts\Install-ScheduledTask.ps1 -Remove"

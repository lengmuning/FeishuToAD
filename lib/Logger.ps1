# Logger.ps1 —— 日志写文件，按天切分
# 需要先 dot-source Common.ps1

$script:LogFilePath = $null
$script:LogSessionId = $null

function Start-SyncLog {
    param(
        [string]$Tag = 'sync'  # sync / dept-preview / single-user / audit
    )
    $root = Get-FeishuToAdRoot
    $logDir = Join-Path $root 'logs'
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $date = Get-Date -Format 'yyyyMMdd'
    $time = Get-Date -Format 'HHmmss'
    $script:LogSessionId = "$date-$time-$Tag"
    $script:LogFilePath = Join-Path $logDir "$date-$Tag.log"
    Write-Log "================ session $script:LogSessionId 开始 ================"
}

function Write-Log {
    param(
        [Parameter(ValueFromPipeline=$true, Position=0)]
        [string]$Message,
        [ValidateSet('INFO','WARN','ERR','OK','ACT','DRY')]
        [string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    if ($script:LogFilePath) {
        Add-Content -Path $script:LogFilePath -Value $line -Encoding UTF8
    }
    # 同时打印到控制台（带颜色）
    switch ($Level) {
        'ERR'  { Write-Host $line -ForegroundColor Red }
        'WARN' { Write-Host $line -ForegroundColor Yellow }
        'OK'   { Write-Host $line -ForegroundColor Green }
        'ACT'  { Write-Host $line -ForegroundColor White }
        'DRY'  { Write-Host $line -ForegroundColor DarkYellow }
        default { Write-Host $line -ForegroundColor Gray }
    }
}

function Remove-OldLogs {
    param([int]$RetentionDays = 30)
    $root = Get-FeishuToAdRoot
    $logDir = Join-Path $root 'logs'
    if (-not (Test-Path $logDir)) { return }
    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -Path $logDir -Filter '*.log' -File |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

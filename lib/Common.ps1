# Common.ps1 —— 通用工具：UTF-8 编码、配置加载、路径
# 所有脚本开头 dot-source 这个文件

# --- UTF-8 编码设置（PowerShell 5.1 中文输出必备）---
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()
$ErrorActionPreference = 'Stop'

# --- 路径解析 ---
function Get-FeishuToAdRoot {
    # 返回 feishutoad/ 根目录绝对路径
    # lib/ 下的脚本：parent 的 parent
    $here = Split-Path -Parent $PSCommandPath
    if ((Split-Path -Leaf $here) -eq 'lib' -or (Split-Path -Leaf $here) -eq 'scripts') {
        return Split-Path -Parent $here
    }
    return $here
}

# --- 配置加载 ---
function Import-SyncConfig {
    param(
        [string]$ConfigPath
    )
    if (-not $ConfigPath) {
        $root = Get-FeishuToAdRoot
        $ConfigPath = Join-Path $root 'config.json'
    }
    if (-not (Test-Path $ConfigPath)) {
        throw "配置文件不存在：$ConfigPath（请复制 config.sample.json 为 config.json 并填值）"
    }
    $raw = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
    $cfg = $raw | ConvertFrom-Json
    # 基本校验
    if (-not $cfg.feishu.appId -or $cfg.feishu.appId -match '^cli_x+$') {
        throw "config.json 未配置 feishu.appId（仍为占位符）"
    }
    if (-not $cfg.feishu.appSecret -or $cfg.feishu.appSecret -match '^<.*>$') {
        throw "config.json 未配置 feishu.appSecret（仍为占位符）"
    }
    if (-not $cfg.ad.syncRootOu) {
        throw "config.json 未配置 ad.syncRootOu"
    }
    if (-not $cfg.ad.archiveOu) {
        throw "config.json 未配置 ad.archiveOu"
    }
    return $cfg
}

# --- 输出美化 ---
function Write-SectionHeader {
    param([string]$Title)
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
}

function Write-Ok    { param($m) Write-Host "  [OK]   $m" -ForegroundColor Green }
function Write-Warn2 { param($m) Write-Host "  [WARN] $m" -ForegroundColor Yellow }
function Write-Err2  { param($m) Write-Host "  [ERR]  $m" -ForegroundColor Red }
function Write-Info  { param($m) Write-Host "  [INFO] $m" -ForegroundColor Gray }
function Write-Act   { param($m) Write-Host "  [ACT]  $m" -ForegroundColor White }
function Write-Dry   { param($m) Write-Host "  [DRY]  $m" -ForegroundColor DarkYellow }

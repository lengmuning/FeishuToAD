# Test-DeptPreview.ps1 —— 阶段 1：预览飞书部门树，不碰 AD
# 不需要 AD 模块，能独立在任何 Windows/PowerShell 跑
# 用途：确认飞书 API 通、Token 有效、部门结构符合预期

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
& (Join-Path $root 'Sync-FeishuToAD.ps1') -Mode Preview

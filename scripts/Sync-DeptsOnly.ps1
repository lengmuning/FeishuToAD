# Sync-DeptsOnly.ps1 —— 阶段 2：只同步部门 OU 结构，不碰用户
# 加 -WhatIf 只演练不真写
# 例：.\scripts\Sync-DeptsOnly.ps1
#     .\scripts\Sync-DeptsOnly.ps1 -WhatIf

param(
    [switch]$WhatIf
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
& (Join-Path $root 'Sync-FeishuToAD.ps1') -Mode DeptsOnly -WhatIf:$WhatIf

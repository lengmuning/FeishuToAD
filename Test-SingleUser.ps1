# Test-SingleUser.ps1 —— 阶段 3：只同步一个指定工号的员工
# 先跑 -WhatIf 看改动，再去掉 -WhatIf 真执行
# 例：.\scripts\Test-SingleUser.ps1 -EmployeeNo <工号> -WhatIf
#     .\scripts\Test-SingleUser.ps1 -EmployeeNo <工号>

param(
    [Parameter(Mandatory)][string]$EmployeeNo,
    [switch]$WhatIf
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
& (Join-Path $root 'Sync-FeishuToAD.ps1') -Mode SingleUser -EmployeeNo $EmployeeNo -WhatIf:$WhatIf

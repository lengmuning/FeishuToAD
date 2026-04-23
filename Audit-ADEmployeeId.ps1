# Audit-ADEmployeeId.ps1 —— 首次全量前的审计工具
#
# 作用：扫描 AD 里已有用户，列出 employeeID 字段是否填了飞书工号。
# 如果大量现有账号 employeeID 为空，首次全量 Full 会把它们当成"AD 里没有的员工"重复创建。
# 建议：按下面输出的结果，先把 AD 里已有账号的 employeeID 手工/批量填回，再跑 Full 同步。
#
# 用法：
#   .\scripts\Audit-ADEmployeeId.ps1
#   .\scripts\Audit-ADEmployeeId.ps1 -SearchBase "OU=YourOU,DC=example,DC=com"
#   .\scripts\Audit-ADEmployeeId.ps1 -OutputCsv ".\ad-audit.csv"
#   .\scripts\Audit-ADEmployeeId.ps1 -JobNumberPattern '^[A-Za-z]{2}\d{2}[A-Za-z]\d{5,}$'
#
# -JobNumberPattern 是你公司工号的正则。默认 '^[A-Za-z0-9]{4,24}$' 比较宽松。
# 如果你的工号有固定格式（如前缀 + 年份 + 字母 + 流水号），改成更严格的正则能少误报。

param(
    [string]$SearchBase,
    [string]$OutputCsv,
    [string]$JobNumberPattern = '^[A-Za-z0-9]{4,24}$',
    [string]$ConfigPath
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
. (Join-Path $root 'lib\Common.ps1')
. (Join-Path $root 'lib\Logger.ps1')
. (Join-Path $root 'lib\AD-Operations.ps1')

Start-SyncLog -Tag 'audit'
Write-SectionHeader "AD employeeID 字段审计"

$config = Import-SyncConfig -ConfigPath $ConfigPath
Initialize-AdModule

if (-not $SearchBase) { $SearchBase = $config.ad.syncRootOu }
Write-Log "扫描范围: $SearchBase"
Write-Log "工号识别正则: $JobNumberPattern"

$users = Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $SearchBase `
    -Properties employeeID, displayName, mail, sAMAccountName, DistinguishedName, userPrincipalName

$total = $users.Count
$withEmpId = @($users | Where-Object { $_.employeeID })
$withoutEmpId = @($users | Where-Object { -not $_.employeeID })

Write-Log "启用用户总数: $total"
Write-Log "已填 employeeID: $($withEmpId.Count)" -Level OK
Write-Log "未填 employeeID: $($withoutEmpId.Count)" -Level WARN

if ($withoutEmpId.Count -gt 0) {
    Write-Host ""
    Write-Host "[WARN] 以下 $($withoutEmpId.Count) 个用户 employeeID 为空，跑 Full 同步前需先回填：" -ForegroundColor Yellow
    $withoutEmpId | Select-Object sAMAccountName, displayName, mail, userPrincipalName, DistinguishedName |
        Format-Table -AutoSize
}

# 检查 UPN/sAMAccountName 是否已符合工号格式，辅助回填
$samLooksLikeJobNumber = @($users | Where-Object { $_.sAMAccountName -match $JobNumberPattern -and -not $_.employeeID })
if ($samLooksLikeJobNumber.Count -gt 0) {
    Write-Host ""
    Write-Host "[HINT] 以下 $($samLooksLikeJobNumber.Count) 个用户 sAMAccountName 匹配工号格式，可考虑将 sAMAccountName 回写到 employeeID：" -ForegroundColor Cyan
    $samLooksLikeJobNumber | Select-Object sAMAccountName, displayName, mail |
        Format-Table -AutoSize
    Write-Host ""
    Write-Host "    批量回填命令示例（先 -WhatIf 演练）：" -ForegroundColor Cyan
    Write-Host "    Get-ADUser -Filter 'Enabled -eq `$true' -SearchBase '$SearchBase' | Where-Object { `$_.sAMAccountName -match '$JobNumberPattern' } | ForEach-Object { Set-ADUser `$_ -EmployeeID `$_.sAMAccountName -WhatIf }" -ForegroundColor DarkCyan
}

if ($OutputCsv) {
    $csvDir = Split-Path -Parent $OutputCsv
    if ($csvDir -and -not (Test-Path $csvDir)) {
        New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
        Write-Log "已创建输出目录: $csvDir" -Level OK
    }
    $users | Select-Object sAMAccountName, employeeID, displayName, mail, userPrincipalName, DistinguishedName |
        Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Log "审计报告已导出: $OutputCsv" -Level OK
}

Write-Log "===== 审计结束 =====" -Level OK

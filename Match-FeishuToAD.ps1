# Match-FeishuToAD.ps1 —— 存量 AD 账号 employeeID 智能回填
#
# 场景：AD 里已有账号，但 employeeID 字段为空，跑 Full 同步前必须先回填。
# 匹配策略（按优先级）：
#   P1: AD sAMAccountName 和飞书工号完全匹配（大小写不敏感）
#   P2: AD mail 和飞书邮箱完全匹配（大小写不敏感）
# 匹配不上的账号会列为"孤儿"，脚本不会动它们，需要人工处理。
#
# 用法：
#   .\scripts\Match-FeishuToAD.ps1                              # dry run，只看报告
#   .\scripts\Match-FeishuToAD.ps1 -Apply                       # 真回填
#   .\scripts\Match-FeishuToAD.ps1 -SearchBase "OU=YourOU,DC=example,DC=com"
#   .\scripts\Match-FeishuToAD.ps1 -OutputCsv ".\match.csv"

[CmdletBinding()]
param(
    [switch]$Apply,
    [string]$SearchBase,
    [string]$OutputCsv,
    [string]$ConfigPath
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
. (Join-Path $root 'lib\Common.ps1')
. (Join-Path $root 'lib\Logger.ps1')
. (Join-Path $root 'lib\Feishu-Api.ps1')
. (Join-Path $root 'lib\AD-Operations.ps1')

Start-SyncLog -Tag 'match'
$modeTag = if ($Apply) { 'APPLY' } else { 'DRY-RUN' }
Write-SectionHeader "飞书 ↔ AD 存量账号智能匹配 [$modeTag]"

$config = Import-SyncConfig -ConfigPath $ConfigPath
Initialize-AdModule

if (-not $SearchBase) { $SearchBase = $config.ad.syncRootOu }
if (-not $OutputCsv) {
    $logDir = Join-Path $root 'logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $OutputCsv = Join-Path $logDir ("match-report-" + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.csv')
}

# 确保 OutputCsv 父目录存在
$csvDir = Split-Path -Parent $OutputCsv
if ($csvDir -and -not (Test-Path $csvDir)) {
    New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
    Write-Log "已创建输出目录: $csvDir" -Level OK
}

Write-Log "AD 扫描范围: $SearchBase"
Write-Log "报告输出:    $OutputCsv"
Write-Log "执行模式:    $modeTag"

# ============================================================
# 1. 拉飞书全量在职员工
# ============================================================
Write-SectionHeader "拉飞书在职员工"
$token = Get-FeishuTenantAccessToken -Config $config
$depts = Get-FeishuAllDepartments -Config $config -Token $token
$deptIds = @($depts | ForEach-Object { $_.OpenId })
Write-Log "飞书 $($deptIds.Count) 个启用部门"

$rawEmps = Get-FeishuEmployeesByDeptIds -Config $config -Token $token -DeptOpenIds $deptIds
$feishuEmps = @($rawEmps | ForEach-Object { Convert-FeishuEmployeeToMapped -Employee $_ })
Write-Log "飞书在职员工总数: $($feishuEmps.Count)" -Level OK

# 建索引
$byJobNumber = @{}
$byEmail = @{}
foreach ($e in $feishuEmps) {
    if ($e.JobNumber) { $byJobNumber[$e.JobNumber.ToLower()] = $e }
    if ($e.Email) { $byEmail[$e.Email.ToLower()] = $e }
}
Write-Log "按工号索引: $($byJobNumber.Count) 条 | 按邮箱索引: $($byEmail.Count) 条"

# ============================================================
# 2. 扫 AD 空 employeeID 启用用户
# ============================================================
Write-SectionHeader "扫 AD 空 employeeID 启用用户"
$adUsers = Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $SearchBase `
    -Properties employeeID, displayName, givenName, sn, mail, sAMAccountName, userPrincipalName, DistinguishedName
$emptyIdUsers = @($adUsers | Where-Object { -not $_.employeeID })
Write-Log "启用用户总数: $($adUsers.Count) | 空 employeeID: $($emptyIdUsers.Count)"

# ============================================================
# 3. 匹配
# ============================================================
Write-SectionHeader "匹配"

$rows = New-Object System.Collections.Generic.List[object]
$p1 = 0; $p2 = 0; $orphan = 0; $ambiguous = 0

foreach ($u in $emptyIdUsers) {
    $hit = $null
    $matchBy = $null

    # P1: sAMAccountName 和工号完全匹配
    if ($u.sAMAccountName) {
        $key = $u.sAMAccountName.ToLower()
        if ($byJobNumber.ContainsKey($key)) {
            $hit = $byJobNumber[$key]
            $matchBy = 'P1-sAMAccountName=JobNumber'
            $p1++
        }
    }

    # P2: mail 和飞书邮箱完全匹配
    if (-not $hit -and $u.mail) {
        $key = $u.mail.ToLower().Trim()
        if ($byEmail.ContainsKey($key)) {
            $hit = $byEmail[$key]
            $matchBy = 'P2-mail=FeishuEmail'
            $p2++
        }
    }

    if ($hit) {
        # 冲突检查：如果飞书员工的工号已经被别的 AD 账号占用（已有 employeeID 的）
        $existingWithId = Get-AdUserByEmployeeId -EmployeeId $hit.JobNumber
        if ($existingWithId -and $existingWithId.DistinguishedName -ne $u.DistinguishedName) {
            $matchBy += ' (CONFLICT: 工号已被其他账号占用)'
            $ambiguous++
            $rows.Add([PSCustomObject]@{
                Status         = 'CONFLICT'
                sAMAccountName = $u.sAMAccountName
                displayName    = $u.displayName
                mail           = $u.mail
                UPN            = $u.userPrincipalName
                FeishuName     = $hit.Name
                FeishuJobNo    = $hit.JobNumber
                FeishuEmail    = $hit.Email
                MatchBy        = $matchBy
                ExistingOwner  = $existingWithId.DistinguishedName
            }) | Out-Null
            continue
        }
        $rows.Add([PSCustomObject]@{
            Status         = 'MATCHED'
            sAMAccountName = $u.sAMAccountName
            displayName    = $u.displayName
            mail           = $u.mail
            UPN            = $u.userPrincipalName
            FeishuName     = $hit.Name
            FeishuJobNo    = $hit.JobNumber
            FeishuEmail    = $hit.Email
            MatchBy        = $matchBy
            ExistingOwner  = ''
        }) | Out-Null
    } else {
        $orphan++
        $rows.Add([PSCustomObject]@{
            Status         = 'ORPHAN'
            sAMAccountName = $u.sAMAccountName
            displayName    = $u.displayName
            mail           = $u.mail
            UPN            = $u.userPrincipalName
            FeishuName     = ''
            FeishuJobNo    = ''
            FeishuEmail    = ''
            MatchBy        = '无匹配（飞书无此人或邮箱不一致）'
            ExistingOwner  = ''
        }) | Out-Null
    }
}

Write-Log "匹配结果: P1(sAM=工号)=$p1  P2(mail=邮箱)=$p2  孤儿=$orphan  冲突=$ambiguous" -Level OK

# ============================================================
# 4. 输出 CSV 报告
# ============================================================
$rows | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Log "CSV 报告已导出: $OutputCsv" -Level OK

# 打印孤儿和冲突详情
$orphans = @($rows | Where-Object { $_.Status -eq 'ORPHAN' })
if ($orphans.Count -gt 0) {
    Write-Host ""
    Write-Host "⚠️ 以下 $($orphans.Count) 个 AD 账号在飞书里匹配不到（需人工处理）：" -ForegroundColor Yellow
    $orphans | Select-Object sAMAccountName, displayName, mail, UPN | Format-Table -AutoSize
}

$conflicts = @($rows | Where-Object { $_.Status -eq 'CONFLICT' })
if ($conflicts.Count -gt 0) {
    Write-Host ""
    Write-Host "❗ 以下 $($conflicts.Count) 个匹配有冲突（飞书工号已被其他 AD 账号占用）：" -ForegroundColor Red
    $conflicts | Select-Object sAMAccountName, displayName, FeishuJobNo, ExistingOwner | Format-Table -AutoSize
}

# ============================================================
# 5. 回填（-Apply）
# ============================================================
$toApply = @($rows | Where-Object { $_.Status -eq 'MATCHED' })
if ($Apply) {
    Write-SectionHeader "开始回填 employeeID（$($toApply.Count) 个账号）"
    $ok = 0; $fail = 0
    foreach ($r in $toApply) {
        try {
            Set-ADUser -Identity $r.sAMAccountName -EmployeeID $r.FeishuJobNo -ErrorAction Stop
            Write-Log "回填 $($r.sAMAccountName) -> employeeID=$($r.FeishuJobNo) ($($r.FeishuName))" -Level OK
            $ok++
        } catch {
            Write-Log "回填失败 $($r.sAMAccountName): $_" -Level ERR
            $fail++
        }
    }
    Write-Log "回填完成: 成功=$ok 失败=$fail" -Level OK
} else {
    Write-Host ""
    Write-Host "============================================================"
    Write-Host "  DRY-RUN 模式，未实际回填。预计将回填 $($toApply.Count) 个账号。" -ForegroundColor Yellow
    Write-Host "  确认 CSV 报告无误后，加 -Apply 真执行：" -ForegroundColor Yellow
    Write-Host "  .\scripts\Match-FeishuToAD.ps1 -SearchBase '$SearchBase' -Apply" -ForegroundColor Cyan
    Write-Host "============================================================"
}

Write-Log "===== match 结束 =====" -Level OK

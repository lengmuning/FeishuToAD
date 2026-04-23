# Sync-FeishuToAD.ps1 —— 主编排脚本
#
# 用法：
#   .\Sync-FeishuToAD.ps1 -Mode Preview              # 预览部门树，不碰 AD
#   .\Sync-FeishuToAD.ps1 -Mode DeptsOnly            # 只同步 OU 结构
#   .\Sync-FeishuToAD.ps1 -Mode DeptsOnly -WhatIf    # 部门同步 dry-run
#   .\Sync-FeishuToAD.ps1 -Mode SingleUser -EmployeeNo <工号>
#   .\Sync-FeishuToAD.ps1 -Mode SingleUser -EmployeeNo <工号> -WhatIf
#   .\Sync-FeishuToAD.ps1 -Mode Full                 # 完整同步（部门 + 全员 + 离职处理）
#   .\Sync-FeishuToAD.ps1 -Mode Full -WhatIf         # 完整 dry-run
#
# 日志：logs/yyyyMMdd-<tag>.log

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Preview','DeptsOnly','SingleUser','Full')]
    [string]$Mode,
    [string]$EmployeeNo,
    [switch]$WhatIf,
    [string]$ConfigPath
)

# --- 加载 lib ---
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'lib\Common.ps1')
. (Join-Path $scriptDir 'lib\Logger.ps1')
. (Join-Path $scriptDir 'lib\Feishu-Api.ps1')
. (Join-Path $scriptDir 'lib\AD-Operations.ps1')

# --- 启动 ---
$tag = switch ($Mode) {
    'Preview'    { 'dept-preview' }
    'DeptsOnly'  { 'depts-only' }
    'SingleUser' { "user-$EmployeeNo" }
    'Full'       { 'full-sync' }
}
Start-SyncLog -Tag $tag

$dryTag = if ($WhatIf) { ' (DRY RUN)' } else { '' }
Write-SectionHeader "Feishu → AD 同步 [$Mode]$dryTag"
$config = Import-SyncConfig -ConfigPath $ConfigPath
Write-Log "配置加载成功 | 同步根 OU: $($config.ad.syncRootOu)"
Write-Log "归档 OU: $($config.ad.archiveOu) | UPN 后缀: $($config.ad.upnSuffix)"

if ($Mode -ne 'Preview') {
    Initialize-AdModule
    # 校验 OU 存在
    try {
        $null = Get-ADOrganizationalUnit -Identity $config.ad.syncRootOu
    } catch {
        throw "同步根 OU 不存在：$($config.ad.syncRootOu)。请先在 AD 里创建，或改 config.json"
    }
    try {
        $null = Get-ADOrganizationalUnit -Identity $config.ad.archiveOu
    } catch {
        throw "归档 OU 不存在：$($config.ad.archiveOu)。请先在 AD 里创建，或改 config.json"
    }
    Write-Log "AD 根 OU 和归档 OU 均存在" -Level OK
}

$token = Get-FeishuTenantAccessToken -Config $config

# =================================================================
# 函数：同步部门 -> OU
# =================================================================
function Invoke-DeptSync {
    param($Config, $Token, [switch]$DryRun, [switch]$PreviewOnly)
    Write-SectionHeader "部门同步"

    $depts = Get-FeishuAllDepartments -Config $Config -Token $Token
    Write-Log "飞书拉到 $($depts.Count) 个启用部门"

    # OpenId -> OU DN 映射表（本次运行维护，供后续员工同步使用）
    $deptOuMap = @{}
    $deptNameMap = @{}   # OpenId -> 部门名

    # 递归/按 BFS 顺序处理（父已在前，见 Feishu-Api 里排序）
    $created = 0; $updated = 0; $moved = 0
    foreach ($d in $depts) {
        $parentDn = if ($d.ParentOpenId -eq '0' -or -not $deptOuMap.ContainsKey($d.ParentOpenId)) {
            $Config.ad.syncRootOu
        } else {
            $deptOuMap[$d.ParentOpenId]
        }

        if ($PreviewOnly) {
            $indentLevel = if ($d.ParentOpenId -eq '0') { 1 } else { 2 }
            $indent = '  ' * $indentLevel
            Write-Host "$indent- $($d.Name)  [feishu:$($d.OpenId)]  -> parent: $parentDn" -ForegroundColor Gray
            $deptOuMap[$d.OpenId] = "OU=$(ConvertTo-SafeOuName $d.Name),$parentDn"
            $deptNameMap[$d.OpenId] = $d.Name
            continue
        }

        # 实际同步
        $existingOu = Get-OuByFeishuId -FeishuDeptId $d.OpenId -SearchBaseOu $Config.ad.syncRootOu
        if (-not $existingOu) {
            # 名字同名兜底（旧 OU 可能没打 description 标签）
            $existingOu = Get-OuByName -OuName (ConvertTo-SafeOuName $d.Name) -SearchBaseOu $parentDn
            if ($existingOu) {
                Write-Log "按名字匹配到已有 OU $($existingOu.DistinguishedName)，回写 feishu:$($d.OpenId) 到 description" -Level WARN
                if (-not $DryRun) {
                    Update-FeishuOu -ExistingOu $existingOu -ExpectedName $d.Name -FeishuDeptId $d.OpenId | Out-Null
                    $updated++
                }
            }
        }

        if (-not $existingOu) {
            # 新建
            $newOu = New-FeishuOu -Name $d.Name -FeishuDeptId $d.OpenId -ParentOu $parentDn -WhatIfMode:$DryRun
            if ($newOu) { $created++ }
            # dry run 时模拟 DN
            $ouDn = if ($newOu) { $newOu.DistinguishedName } else { "OU=$(ConvertTo-SafeOuName $d.Name),$parentDn" }
        } else {
            # 移动
            if (Move-OuToNewParent -ExistingOu $existingOu -NewParentOu $parentDn -WhatIfMode:$DryRun) {
                $moved++
                # move 后刷新 DN
                if (-not $DryRun) { $existingOu = Get-OuByFeishuId -FeishuDeptId $d.OpenId -SearchBaseOu $Config.ad.syncRootOu }
            }
            # 改名/description
            if (Update-FeishuOu -ExistingOu $existingOu -ExpectedName $d.Name -FeishuDeptId $d.OpenId -WhatIfMode:$DryRun) {
                $updated++
                if (-not $DryRun) { $existingOu = Get-OuByFeishuId -FeishuDeptId $d.OpenId -SearchBaseOu $Config.ad.syncRootOu }
            }
            $ouDn = $existingOu.DistinguishedName
        }

        $deptOuMap[$d.OpenId] = $ouDn
        $deptNameMap[$d.OpenId] = $d.Name
    }

    Write-Log "部门同步完成: 创建=$created 更新=$updated 移动=$moved 总计=$($depts.Count)" -Level OK
    return @{ DeptOuMap = $deptOuMap; DeptNameMap = $deptNameMap; Depts = $depts }
}

# =================================================================
# 函数：同步员工
# =================================================================
function Invoke-UserSync {
    param(
        $Config, $Token,
        [hashtable]$DeptOuMap,
        [hashtable]$DeptNameMap,
        [switch]$DryRun,
        [string]$OnlyEmployeeNo  # 单人模式
    )
    $hdr = "员工同步"
    if ($OnlyEmployeeNo) { $hdr = "员工同步 - 单人 $OnlyEmployeeNo" }
    Write-SectionHeader $hdr

    $allDeptIds = @($DeptOuMap.Keys)
    if ($allDeptIds.Count -eq 0) {
        Write-Log "没有可同步的部门，跳过员工同步" -Level WARN
        return
    }

    Write-Log "从 $($allDeptIds.Count) 个飞书部门拉在职员工..."
    $employees = Get-FeishuEmployeesByDeptIds -Config $Config -Token $Token -DeptOpenIds $allDeptIds
    Write-Log "飞书拉到 $($employees.Count) 个在职员工"

    $created = 0; $updated = 0; $skipped = 0; $errors = 0
    $feishuActiveJobNumbers = New-Object System.Collections.Generic.HashSet[string]

    foreach ($emp in $employees) {
        $m = Convert-FeishuEmployeeToMapped -Employee $emp

        # 单人模式过滤
        if ($OnlyEmployeeNo -and $m.JobNumber -ne $OnlyEmployeeNo) { continue }

        if (-not (Test-ValidJobNumber $m.JobNumber)) {
            Write-Log "跳过: 工号无效 '$($m.JobNumber)' (name=$($m.Name))" -Level WARN
            $skipped++; continue
        }
        if (-not $m.Name) {
            Write-Log "跳过: 姓名为空 (工号=$($m.JobNumber))" -Level WARN
            $skipped++; continue
        }

        $feishuActiveJobNumbers.Add($m.JobNumber) | Out-Null

        # 定位目标 OU（取第一个部门）
        $targetOu = $null
        $targetDeptName = $null
        foreach ($did in $m.DeptOpenIds) {
            if ($DeptOuMap.ContainsKey($did)) {
                $targetOu = $DeptOuMap[$did]
                $targetDeptName = $DeptNameMap[$did]
                break
            }
        }
        if (-not $targetOu) {
            Write-Log "跳过: 找不到员工 $($m.JobNumber)($($m.Name)) 的部门 OU (飞书部门 ids=$($m.DeptOpenIds -join ','))" -Level WARN
            $skipped++; continue
        }

        try {
            $existing = Get-AdUserByEmployeeId -EmployeeId $m.JobNumber
            if ($existing) {
                # 存量更新（只改属性，不改密码，不改 Enabled）
                $changes = Update-FeishuAdUser -Config $Config -ExistingUser $existing `
                    -JobNumber $m.JobNumber -Name $m.Name -Email $m.Email `
                    -TargetOu $targetOu -TargetDeptName $targetDeptName -WhatIfMode:$DryRun
                if ($changes.Count -gt 0) { $updated++ } else { Write-Log "无变化: $($m.JobNumber) ($($m.Name))" }
            } else {
                # 新建
                $null = New-FeishuAdUser -Config $Config -JobNumber $m.JobNumber -Name $m.Name `
                    -Email $m.Email -TargetOu $targetOu -WhatIfMode:$DryRun
                $created++
            }
        } catch {
            Write-Log "处理员工 $($m.JobNumber) 失败: $_" -Level ERR
            $errors++
        }
    }

    Write-Log "员工同步完成: 新建=$created 更新=$updated 跳过=$skipped 失败=$errors" -Level OK

    # 离职处理：仅完整模式才做
    if (-not $OnlyEmployeeNo) {
        Invoke-DepartedUserArchive -Config $Config -FeishuActiveJobNumbers $feishuActiveJobNumbers -DryRun:$DryRun
    }
}

# =================================================================
# 函数：离职差集处理
# =================================================================
function Invoke-DepartedUserArchive {
    param($Config, [System.Collections.Generic.HashSet[string]]$FeishuActiveJobNumbers, [switch]$DryRun)
    Write-SectionHeader "离职差集处理"
    $adUsers = Get-AllActiveUsersUnderOu -SyncRootOu $Config.ad.syncRootOu
    Write-Log "同步根 OU 下启用用户数: $($adUsers.Count)"
    $archived = 0
    foreach ($u in $adUsers) {
        if (-not $u.employeeID) { continue }
        if ($FeishuActiveJobNumbers.Contains($u.employeeID)) { continue }
        # 飞书已无该工号对应在职员工 -> 归档
        try {
            Disable-AndArchiveAdUser -ExistingUser $u -ArchiveOu $Config.ad.archiveOu -WhatIfMode:$DryRun | Out-Null
            $archived++
        } catch {
            Write-Log "归档用户 $($u.sAMAccountName) 失败: $_" -Level ERR
        }
    }
    Write-Log "离职归档完成: 共处理 $archived 个用户" -Level OK
}

# =================================================================
# 入口分发
# =================================================================
switch ($Mode) {
    'Preview' {
        Invoke-DeptSync -Config $config -Token $token -DryRun -PreviewOnly | Out-Null
        Write-Log "预览完成（未写 AD）" -Level OK
    }
    'DeptsOnly' {
        Invoke-DeptSync -Config $config -Token $token -DryRun:$WhatIf | Out-Null
    }
    'SingleUser' {
        if (-not $EmployeeNo) { throw "-Mode SingleUser 必须同时传 -EmployeeNo <工号>" }
        # 单人模式也需要先把部门 OU 建好（否则找不到目标 OU）
        $deptResult = Invoke-DeptSync -Config $config -Token $token -DryRun:$WhatIf
        Invoke-UserSync -Config $config -Token $token `
            -DeptOuMap $deptResult.DeptOuMap -DeptNameMap $deptResult.DeptNameMap `
            -DryRun:$WhatIf -OnlyEmployeeNo $EmployeeNo
    }
    'Full' {
        $deptResult = Invoke-DeptSync -Config $config -Token $token -DryRun:$WhatIf
        Invoke-UserSync -Config $config -Token $token `
            -DeptOuMap $deptResult.DeptOuMap -DeptNameMap $deptResult.DeptNameMap `
            -DryRun:$WhatIf
    }
}

Remove-OldLogs -RetentionDays ([int]$config.sync.logRetentionDays)
Write-Log "===== session $tag 结束 =====" -Level OK

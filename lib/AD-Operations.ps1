# AD-Operations.ps1 —— AD 操作封装
# 需要先 dot-source Common.ps1 和 Logger.ps1
# 运行环境：加域的 Windows Server + RSAT (ActiveDirectory 模块)

# ----------- 模块初始化 -----------
function Initialize-AdModule {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "未安装 ActiveDirectory PowerShell 模块。请以管理员身份运行：Install-WindowsFeature RSAT-AD-PowerShell"
    }
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log "ActiveDirectory 模块已加载"
}

# ----------- OU 操作 -----------
function Get-OuByFeishuId {
    # 按飞书 department_id（存在 description 字段里）查 OU
    param(
        [Parameter(Mandatory)][string]$FeishuDeptId,
        [Parameter(Mandatory)][string]$SearchBaseOu
    )
    $filter = "description -eq 'feishu:$FeishuDeptId'"
    $ou = Get-ADOrganizationalUnit -Filter $filter -SearchBase $SearchBaseOu -ErrorAction SilentlyContinue
    return $ou
}

function Get-OuByName {
    # 按 OU 名在 SearchBase 下一层查
    param(
        [Parameter(Mandatory)][string]$OuName,
        [Parameter(Mandatory)][string]$SearchBaseOu
    )
    $escapedName = $OuName -replace "'", "''"
    $filter = "Name -eq '$escapedName'"
    $ou = Get-ADOrganizationalUnit -Filter $filter -SearchBase $SearchBaseOu -SearchScope OneLevel -ErrorAction SilentlyContinue
    return $ou
}

function New-FeishuOu {
    # 在指定父 OU 下创建新 OU，description 标记飞书 id
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$FeishuDeptId,
        [Parameter(Mandatory)][string]$ParentOu,
        [switch]$WhatIfMode
    )
    $cleanName = ConvertTo-SafeOuName $Name
    if ($WhatIfMode) {
        Write-Log "[DRY] 将创建 OU: CN=$cleanName,$ParentOu  (feishu:$FeishuDeptId)" -Level DRY
        return $null
    }
    try {
        New-ADOrganizationalUnit -Name $cleanName `
            -Path $ParentOu `
            -Description "feishu:$FeishuDeptId" `
            -ProtectedFromAccidentalDeletion $false `
            -ErrorAction Stop
        $newOu = Get-OuByFeishuId -FeishuDeptId $FeishuDeptId -SearchBaseOu $ParentOu
        Write-Log "创建 OU: $($newOu.DistinguishedName)" -Level OK
        return $newOu
    } catch {
        Write-Log "创建 OU 失败 Name=$cleanName Parent=${ParentOu}: $_" -Level ERR
        throw
    }
}

function Update-FeishuOu {
    # 同步已有 OU 的名字/description
    param(
        [Parameter(Mandatory)]$ExistingOu,
        [Parameter(Mandatory)][string]$ExpectedName,
        [Parameter(Mandatory)][string]$FeishuDeptId,
        [switch]$WhatIfMode
    )
    $cleanName = ConvertTo-SafeOuName $ExpectedName
    $changed = $false
    if ($ExistingOu.Name -ne $cleanName) {
        if ($WhatIfMode) {
            Write-Log "[DRY] 将改名 OU: $($ExistingOu.Name) -> $cleanName" -Level DRY
        } else {
            Rename-ADObject -Identity $ExistingOu.DistinguishedName -NewName $cleanName -ErrorAction Stop
            Write-Log "改名 OU: $($ExistingOu.Name) -> $cleanName" -Level OK
            $changed = $true
        }
    }
    $expectedDesc = "feishu:$FeishuDeptId"
    if ($ExistingOu.Description -ne $expectedDesc) {
        if ($WhatIfMode) {
            Write-Log "[DRY] 将更新 OU description: $($ExistingOu.Name) -> $expectedDesc" -Level DRY
        } else {
            Set-ADOrganizationalUnit -Identity $ExistingOu.DistinguishedName -Description $expectedDesc -ErrorAction Stop
            $changed = $true
        }
    }
    return $changed
}

function Move-OuToNewParent {
    param(
        [Parameter(Mandatory)]$ExistingOu,
        [Parameter(Mandatory)][string]$NewParentOu,
        [switch]$WhatIfMode
    )
    $currentParent = ($ExistingOu.DistinguishedName -split ',', 2)[1]
    if ($currentParent -eq $NewParentOu) { return $false }
    if ($WhatIfMode) {
        Write-Log "[DRY] 将移动 OU: $($ExistingOu.DistinguishedName) -> $NewParentOu" -Level DRY
        return $true
    }
    Move-ADObject -Identity $ExistingOu.DistinguishedName -TargetPath $NewParentOu -ErrorAction Stop
    Write-Log "移动 OU: $($ExistingOu.Name) -> $NewParentOu" -Level OK
    return $true
}

function ConvertTo-SafeOuName {
    # OU 名过滤非法字符
    param([string]$Name)
    if (-not $Name) { return '未命名部门' }
    $cleaned = $Name -replace '[,\\#+<>;"=/]', '_'
    $cleaned = $cleaned.Trim()
    if ($cleaned.Length -gt 60) { $cleaned = $cleaned.Substring(0, 60) }
    if (-not $cleaned) { return '未命名部门' }
    return $cleaned
}

# ----------- 用户操作 -----------
function Get-AdUserByEmployeeId {
    # 按 employeeID 精确匹配（全域搜索）
    param([Parameter(Mandatory)][string]$EmployeeId)
    $escaped = $EmployeeId -replace "'", "''"
    $u = Get-ADUser -Filter "employeeID -eq '$escaped'" `
        -Properties employeeID, displayName, givenName, sn, mail, department, userPrincipalName, sAMAccountName, Enabled, DistinguishedName `
        -ErrorAction SilentlyContinue
    return $u
}

function Get-AdUserBySamAccount {
    param([Parameter(Mandatory)][string]$SamAccountName)
    try {
        return Get-ADUser -Identity $SamAccountName `
            -Properties employeeID, displayName, givenName, sn, mail, department, userPrincipalName, sAMAccountName, Enabled, DistinguishedName `
            -ErrorAction Stop
    } catch {
        return $null
    }
}

function New-FeishuAdUser {
    # 创建新 AD 用户
    # 规则：sAMAccountName=工号，UPN=工号@UpnSuffix，CN=姓名（冲突时 fallback 姓名+工号），sn=姓名全名，givenName 空
    param(
        [Parameter(Mandatory)][object]$Config,
        [Parameter(Mandatory)][string]$JobNumber,
        [Parameter(Mandatory)][string]$Name,
        [string]$Email,
        [Parameter(Mandatory)][string]$TargetOu,
        [switch]$WhatIfMode
    )
    $sam = $JobNumber
    # sAMAccountName 硬限制 20 字符
    if ($sam.Length -gt 20) { $sam = $sam.Substring(0, 20) }

    $upnSuffix = $Config.ad.upnSuffix
    if (-not $upnSuffix.StartsWith('@')) { $upnSuffix = '@' + $upnSuffix }
    $upn = "$JobNumber$upnSuffix"

    # CN 冲突检测
    $cn = $Name
    $existing = Get-ADUser -Filter "Name -eq '$cn'" -SearchBase $TargetOu -SearchScope OneLevel -ErrorAction SilentlyContinue
    if ($existing) {
        $cn = "$Name ($JobNumber)"
        Write-Log "CN 冲突，使用 fallback CN: $cn" -Level WARN
    }

    if ($WhatIfMode) {
        Write-Log "[DRY] 将创建用户: sAMAccountName=$sam UPN=$upn CN=$cn DisplayName=$Name sn=$Name mail=$Email OU=$TargetOu" -Level DRY
        return $null
    }

    $securePwd = ConvertTo-SecureString -String $Config.user.defaultPassword -AsPlainText -Force
    $params = @{
        Name                  = $cn
        SamAccountName        = $sam
        UserPrincipalName     = $upn
        DisplayName           = $Name
        Surname               = $Name    # 照你现有约定：整个中文姓名放 sn
        EmployeeID            = $JobNumber
        Path                  = $TargetOu
        AccountPassword       = $securePwd
        Enabled               = [bool]$Config.user.enabledOnCreate
        ChangePasswordAtLogon = [bool]$Config.user.changePasswordAtLogon
        PasswordNeverExpires  = [bool]$Config.user.passwordNeverExpires
        ErrorAction           = 'Stop'
    }
    if ($Email) { $params.EmailAddress = $Email }

    New-ADUser @params
    $newUser = Get-AdUserByEmployeeId -EmployeeId $JobNumber
    Write-Log "创建用户: $($newUser.DistinguishedName) (mail=$Email, UPN=$upn)" -Level OK
    return $newUser
}

function Update-FeishuAdUser {
    # 更新已有用户（姓名/邮箱/department 字段 + OU 位置）
    # 【绝不修改密码、绝不修改 Enabled 状态（除非从归档恢复）】
    param(
        [Parameter(Mandatory)][object]$Config,
        [Parameter(Mandatory)]$ExistingUser,
        [Parameter(Mandatory)][string]$JobNumber,
        [Parameter(Mandatory)][string]$Name,
        [string]$Email,
        [Parameter(Mandatory)][string]$TargetOu,
        [Parameter(Mandatory)][string]$TargetDeptName,
        [switch]$WhatIfMode
    )
    $changes = @{}
    $setParams = @{}

    # DisplayName 纠正
    if ($ExistingUser.DisplayName -ne $Name) {
        $changes['DisplayName'] = "$($ExistingUser.DisplayName) -> $Name"
        $setParams['DisplayName'] = $Name
    }
    # Surname 纠正（中文姓名整体放 sn）
    if ($ExistingUser.Surname -ne $Name) {
        $changes['Surname'] = "$($ExistingUser.Surname) -> $Name"
        $setParams['Surname'] = $Name
    }
    # 邮箱纠正
    if ($Email -and $ExistingUser.mail -ne $Email) {
        $changes['Mail'] = "$($ExistingUser.mail) -> $Email"
        $setParams['EmailAddress'] = $Email
    }
    # department 字段（AD 属性，不是 OU 位置）
    if ($ExistingUser.department -ne $TargetDeptName) {
        $changes['Department'] = "$($ExistingUser.department) -> $TargetDeptName"
        $setParams['Department'] = $TargetDeptName
    }
    # employeeID 兜底（理论上应该已经对上了，因为我们就是按它查的）
    if ($ExistingUser.employeeID -ne $JobNumber) {
        $changes['EmployeeID'] = "$($ExistingUser.employeeID) -> $JobNumber"
        $setParams['EmployeeID'] = $JobNumber
    }

    # 属性更新
    if ($setParams.Count -gt 0) {
        if ($WhatIfMode) {
            Write-Log "[DRY] 将更新用户 $($ExistingUser.SamAccountName): $($changes | ConvertTo-Json -Compress)" -Level DRY
        } else {
            Set-ADUser -Identity $ExistingUser.DistinguishedName @setParams -ErrorAction Stop
            Write-Log "更新用户 $($ExistingUser.SamAccountName): $($changes | ConvertTo-Json -Compress)" -Level OK
        }
    }

    # CN 重命名（如果姓名变了，CN 也跟着改）
    $currentCn = ($ExistingUser.DistinguishedName -split ',', 2)[0] -replace '^CN=', ''
    $expectedCn = $Name
    if ($currentCn -ne $Name -and $currentCn -ne "$Name ($JobNumber)") {
        if ($WhatIfMode) {
            Write-Log "[DRY] 将改 CN: $currentCn -> $expectedCn" -Level DRY
        } else {
            try {
                Rename-ADObject -Identity $ExistingUser.DistinguishedName -NewName $expectedCn -ErrorAction Stop
                # 刷新 DN（rename 后 DN 变了）
                $ExistingUser = Get-AdUserByEmployeeId -EmployeeId $JobNumber
                Write-Log "改名 CN: $currentCn -> $expectedCn ($($ExistingUser.DistinguishedName))" -Level OK
            } catch {
                # 冲突则 fallback
                $fallback = "$Name ($JobNumber)"
                try {
                    Rename-ADObject -Identity $ExistingUser.DistinguishedName -NewName $fallback -ErrorAction Stop
                    $ExistingUser = Get-AdUserByEmployeeId -EmployeeId $JobNumber
                    Write-Log "CN 冲突，已用 fallback: $currentCn -> $fallback" -Level WARN
                } catch {
                    Write-Log "CN 改名失败: $_" -Level ERR
                }
            }
        }
    }

    # OU 位置调整
    $currentParent = ($ExistingUser.DistinguishedName -split ',', 2)[1]
    if ($currentParent -ne $TargetOu) {
        if ($WhatIfMode) {
            Write-Log "[DRY] 将移动用户 $($ExistingUser.SamAccountName): $currentParent -> $TargetOu" -Level DRY
        } else {
            Move-ADObject -Identity $ExistingUser.DistinguishedName -TargetPath $TargetOu -ErrorAction Stop
            Write-Log "移动用户 $($ExistingUser.SamAccountName) -> $TargetOu" -Level OK
        }
    }

    return $changes
}

function Disable-AndArchiveAdUser {
    # 禁用 + 挪到归档 OU（飞书离职处理）
    # 【绝不修改密码、绝不删除对象】
    param(
        [Parameter(Mandatory)]$ExistingUser,
        [Parameter(Mandatory)][string]$ArchiveOu,
        [switch]$WhatIfMode
    )
    $actions = New-Object System.Collections.Generic.List[string]
    if ($ExistingUser.Enabled) {
        if ($WhatIfMode) {
            $actions.Add("disable") | Out-Null
            Write-Log "[DRY] 将禁用用户 $($ExistingUser.SamAccountName)" -Level DRY
        } else {
            Disable-ADAccount -Identity $ExistingUser.DistinguishedName -ErrorAction Stop
            $actions.Add("disabled") | Out-Null
            Write-Log "禁用用户 $($ExistingUser.SamAccountName)" -Level OK
        }
    }
    $currentParent = ($ExistingUser.DistinguishedName -split ',', 2)[1]
    if ($currentParent -ne $ArchiveOu) {
        if ($WhatIfMode) {
            $actions.Add("move to archive") | Out-Null
            Write-Log "[DRY] 将归档用户 $($ExistingUser.SamAccountName) -> $ArchiveOu" -Level DRY
        } else {
            Move-ADObject -Identity $ExistingUser.DistinguishedName -TargetPath $ArchiveOu -ErrorAction Stop
            $actions.Add("archived") | Out-Null
            Write-Log "归档用户 $($ExistingUser.SamAccountName) -> $ArchiveOu" -Level OK
        }
    }
    return $actions
}

function Get-AllActiveUsersUnderOu {
    # 扫描同步根 OU 下所有**启用**用户（用于离职差集计算）
    param([Parameter(Mandatory)][string]$SyncRootOu)
    Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $SyncRootOu `
        -Properties employeeID, displayName, mail, sAMAccountName, DistinguishedName |
        Where-Object { $_.employeeID }
}

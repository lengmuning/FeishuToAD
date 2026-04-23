# Feishu-Api.ps1 —— 飞书 Directory v1 API 封装
# 基于 src/index.js 的逻辑移植到 PowerShell
# 需要先 dot-source Common.ps1 和 Logger.ps1

# ----------- 认证 -----------
function Get-FeishuTenantAccessToken {
    param(
        [Parameter(Mandatory)][object]$Config
    )
    $url = "$($Config.feishu.apiBase)/open-apis/auth/v3/tenant_access_token/internal"
    $body = @{
        app_id     = $Config.feishu.appId
        app_secret = $Config.feishu.appSecret
    } | ConvertTo-Json
    $resp = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType 'application/json; charset=utf-8'
    if ($resp.code -ne 0) {
        throw "飞书 token 获取失败：$($resp | ConvertTo-Json -Depth 5)"
    }
    Write-Log "飞书 tenant_access_token 获取成功（有效期 $($resp.expire) 秒）"
    return $resp.tenant_access_token
}

# ----------- 部门 -----------
function Get-FeishuAllDepartments {
    # 单次过滤拉全量启用部门，客户端 BFS 排序（父在子前）
    param(
        [Parameter(Mandatory)][object]$Config,
        [Parameter(Mandatory)][string]$Token
    )
    $url = "$($Config.feishu.apiBase)/open-apis/directory/v1/departments/filter?department_id_type=open_department_id"
    $headers = @{
        'Authorization' = "Bearer $Token"
        'Content-Type'  = 'application/json; charset=utf-8'
    }
    $pageToken = ''
    $all = New-Object System.Collections.Generic.List[object]
    while ($true) {
        $bodyObj = @{
            filter = @{
                conditions = @(
                    @{ field = 'enabled_status'; operator = 'eq'; value = 'true' }
                )
            }
            required_fields = @('name','parent_department_id','department_id','enabled_status')
            page_request    = @{ page_size = 100 }
        }
        if ($pageToken) { $bodyObj.page_request.page_token = $pageToken }
        $body = $bodyObj | ConvertTo-Json -Depth 10
        $resp = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        if ($resp.code -ne 0) {
            throw "飞书 departments/filter 失败：$($resp | ConvertTo-Json -Depth 5)"
        }
        foreach ($d in $resp.data.departments) { $all.Add($d) | Out-Null }
        if (-not $resp.data.page_response.has_more -or -not $resp.data.page_response.page_token) { break }
        $pageToken = $resp.data.page_response.page_token
    }
    Write-Log "飞书拉到 $($all.Count) 个启用部门"

    # 客户端 BFS 排序：父在子前
    $byOpenId = @{}
    $parentOf = @{}
    foreach ($d in $all) {
        $openId = Get-FeishuIdValue $d.department_id
        if (-not $openId) { continue }
        $byOpenId[$openId] = $d
        $parentId = Get-FeishuIdValue $d.parent_department_id
        if (-not $parentId) { $parentId = '0' }
        $parentOf[$openId] = $parentId
    }
    $childrenOf = @{}
    foreach ($k in $parentOf.Keys) {
        $p = $parentOf[$k]
        if (-not $childrenOf.ContainsKey($p)) { $childrenOf[$p] = New-Object System.Collections.Generic.List[string] }
        $childrenOf[$p].Add($k)
    }
    $ordered = New-Object System.Collections.Generic.List[object]
    $visited = @{}
    $queue = New-Object System.Collections.Generic.Queue[string]
    if ($childrenOf.ContainsKey('0')) {
        foreach ($c in $childrenOf['0']) { $queue.Enqueue($c) }
    }
    while ($queue.Count -gt 0) {
        $openId = $queue.Dequeue()
        if ($visited.ContainsKey($openId)) { continue }
        $visited[$openId] = $true
        $d = $byOpenId[$openId]
        if (-not $d) { continue }
        $ordered.Add([PSCustomObject]@{
            OpenId       = $openId
            ParentOpenId = $parentOf[$openId]
            Name         = $d.name.default_value
            Raw          = $d
        }) | Out-Null
        if ($childrenOf.ContainsKey($openId)) {
            foreach ($c in $childrenOf[$openId]) { $queue.Enqueue($c) }
        }
    }
    # 补漏：找不到父链的孤儿部门也加进来（挂根）
    foreach ($k in $byOpenId.Keys) {
        if ($visited.ContainsKey($k)) { continue }
        $d = $byOpenId[$k]
        $ordered.Add([PSCustomObject]@{
            OpenId       = $k
            ParentOpenId = '0'
            Name         = $d.name.default_value
            Raw          = $d
        }) | Out-Null
    }
    return $ordered
}

function Get-FeishuIdValue {
    # 飞书有些 id 字段是对象 {open_department_id: "..."}，有些直接是字符串
    param($field)
    if ($null -eq $field) { return $null }
    if ($field -is [string]) { return $field }
    if ($field.open_department_id) { return $field.open_department_id }
    if ($field.department_id) { return $field.department_id }
    return $null
}

# ----------- 员工 -----------
function Get-FeishuEmployeesByDept {
    # 按单部门拉在职员工
    param(
        [Parameter(Mandatory)][object]$Config,
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$DeptOpenId
    )
    $url = "$($Config.feishu.apiBase)/open-apis/directory/v1/employees/filter?department_id_type=open_department_id"
    $headers = @{
        'Authorization' = "Bearer $Token"
        'Content-Type'  = 'application/json; charset=utf-8'
    }
    $pageToken = ''
    $all = New-Object System.Collections.Generic.List[object]
    while ($true) {
        $bodyObj = @{
            filter = @{
                conditions = @(
                    @{ field = 'base_info.departments.department_id'; operator = 'eq'; value = $DeptOpenId },
                    @{ field = 'work_info.staff_status'; operator = 'eq'; value = '1' }
                )
            }
            required_fields = @(
                'base_info.name.name',
                'base_info.email',
                'base_info.departments',
                'work_info.job_number',
                'work_info.email',
                'work_info.staff_status'
            )
            page_request = @{ page_size = 100 }
        }
        if ($pageToken) { $bodyObj.page_request.page_token = $pageToken }
        $body = $bodyObj | ConvertTo-Json -Depth 10
        $resp = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        if ($resp.code -ne 0) {
            throw "飞书 employees/filter 失败 (dept=$DeptOpenId)：$($resp | ConvertTo-Json -Depth 5)"
        }
        foreach ($e in $resp.data.employees) { $all.Add($e) | Out-Null }
        if (-not $resp.data.page_response.has_more -or -not $resp.data.page_response.page_token) { break }
        $pageToken = $resp.data.page_response.page_token
    }
    return $all
}

function Get-FeishuEmployeesByDeptIds {
    # 批量按多个部门拉在职员工（最多 50 个 id/批）
    param(
        [Parameter(Mandatory)][object]$Config,
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string[]]$DeptOpenIds
    )
    $url = "$($Config.feishu.apiBase)/open-apis/directory/v1/employees/filter?department_id_type=open_department_id"
    $headers = @{
        'Authorization' = "Bearer $Token"
        'Content-Type'  = 'application/json; charset=utf-8'
    }
    $collected = New-Object System.Collections.Generic.List[object]
    $seenIds = @{}
    $batchSize = 50
    for ($i = 0; $i -lt $DeptOpenIds.Count; $i += $batchSize) {
        $end = [Math]::Min($i + $batchSize - 1, $DeptOpenIds.Count - 1)
        $batch = $DeptOpenIds[$i..$end]
        $pageToken = ''
        while ($true) {
            $bodyObj = @{
                filter = @{
                    conditions = @(
                        @{ field = 'base_info.departments.department_id'; operator = 'in'; value = ($batch | ConvertTo-Json -Compress) },
                        @{ field = 'work_info.staff_status'; operator = 'eq'; value = '1' }
                    )
                }
                required_fields = @(
                    'base_info.name.name',
                    'base_info.email',
                    'base_info.departments',
                    'work_info.job_number',
                    'work_info.email',
                    'work_info.staff_status'
                )
                page_request = @{ page_size = 100 }
            }
            if ($pageToken) { $bodyObj.page_request.page_token = $pageToken }
            $body = $bodyObj | ConvertTo-Json -Depth 10
            $resp = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
            if ($resp.code -ne 0) {
                throw "飞书 employees/filter 批量失败：$($resp | ConvertTo-Json -Depth 5)"
            }
            foreach ($e in $resp.data.employees) {
                $empId = Get-FeishuIdValue $e.employee_id
                if (-not $empId) {
                    $jn = [string]$e.work_info.job_number
                    $nm = [string]$e.base_info.name.name
                    if ($jn) { $empId = "jn:$jn" }
                    elseif ($nm) { $empId = "nm:$nm" }
                    else { $empId = [guid]::NewGuid().ToString() }
                }
                if ($seenIds.ContainsKey($empId)) { continue }
                $seenIds[$empId] = $true
                $collected.Add($e) | Out-Null
            }
            if (-not $resp.data.page_response.has_more -or -not $resp.data.page_response.page_token) { break }
            $pageToken = $resp.data.page_response.page_token
        }
    }
    return $collected
}

# ----------- 字段映射 -----------
function Convert-FeishuEmployeeToMapped {
    # 把飞书原始员工对象转成标准化对象，后续直接喂给 AD 操作
    param([Parameter(Mandatory)]$Employee)

    # 飞书 v1 真实字段路径（经 Inspect-FeishuScope 诊断确认）：
    #   姓名 → base_info.name.name.default_value
    #   工号 → work_info.job_number
    #   邮箱 → base_info.email  (对应 scope: directory:employee.base.email:read)
    $name = $null
    if ($Employee.base_info.name.name.default_value) { $name = [string]$Employee.base_info.name.name.default_value }
    elseif ($Employee.base_info.name.default_value)  { $name = [string]$Employee.base_info.name.default_value }

    $jobNumber = [string]$Employee.work_info.job_number

    $email = $null
    if ($Employee.base_info.email)      { $email = [string]$Employee.base_info.email }
    elseif ($Employee.work_info.email)  { $email = [string]$Employee.work_info.email }

    $deptIds = @()
    if ($Employee.base_info.departments) {
        foreach ($d in $Employee.base_info.departments) {
            $did = Get-FeishuIdValue $d.department_id
            if ($did) { $deptIds += $did }
        }
    }

    return [PSCustomObject]@{
        Name            = $name
        JobNumber       = $jobNumber
        Email           = $email
        DeptOpenIds     = $deptIds
        PrimaryDeptId   = ($deptIds | Select-Object -First 1)
        Raw             = $Employee
    }
}

function Test-ValidJobNumber {
    # 工号必须为 1-24 位字母数字
    param([string]$JobNumber)
    if ([string]::IsNullOrWhiteSpace($JobNumber)) { return $false }
    return $JobNumber -match '^[A-Za-z0-9]{1,24}$'
}

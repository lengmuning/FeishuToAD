# Inspect-FeishuScope.ps1 —— 飞书权限 scope 针对性诊断
# 用不同 required_fields 组合发请求，观察哪些字段真的有返回。
# 如果 email/staff_status 全是 null/missing，基本可以确认是 scope 没生效。

param(
    [string]$JobNumber,
    [string]$ConfigPath
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
. (Join-Path $root 'lib\Common.ps1')
. (Join-Path $root 'lib\Logger.ps1')
. (Join-Path $root 'lib\Feishu-Api.ps1')

Start-SyncLog -Tag 'scope'
Write-SectionHeader "飞书 scope 针对性诊断"

$config = Import-SyncConfig -ConfigPath $ConfigPath
$token = Get-FeishuTenantAccessToken -Config $config

$url = "$($config.feishu.apiBase)/open-apis/directory/v1/employees/filter?department_id_type=open_department_id"
$headers = @{
    'Authorization' = "Bearer $token"
    'Content-Type'  = 'application/json; charset=utf-8'
}

# 逐个测试不同字段，只查 1 个工号的员工
$testCases = @(
    @{ Name = '仅 job_number (对照组)';     Fields = @('work_info.job_number') },
    @{ Name = '+ base_info.email';          Fields = @('work_info.job_number','base_info.email') },
    @{ Name = '+ work_info.email';          Fields = @('work_info.job_number','work_info.email') },
    @{ Name = '+ work_info.work_email';     Fields = @('work_info.job_number','work_info.work_email') },
    @{ Name = '+ work_info.staff_status';   Fields = @('work_info.job_number','work_info.staff_status') },
    @{ Name = '+ base_info.mobile';         Fields = @('work_info.job_number','base_info.mobile') },
    @{ Name = '+ base_info.name.name';      Fields = @('work_info.job_number','base_info.name.name') }
)

foreach ($tc in $testCases) {
    Write-Host ""
    Write-Host "────────────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host "测试: $($tc.Name)" -ForegroundColor Cyan
    Write-Host "字段: $($tc.Fields -join ', ')" -ForegroundColor DarkCyan

    $bodyObj = @{
        filter = @{
            conditions = @(
                @{ field = 'work_info.staff_status'; operator = 'eq'; value = '1' }
            )
        }
        required_fields = $tc.Fields
        page_request    = @{ page_size = 3 }
    }
    $body = $bodyObj | ConvertTo-Json -Depth 10

    try {
        $resp = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        if ($resp.code -ne 0) {
            Write-Host "  ❌ API 返回非 0: code=$($resp.code) msg=$($resp.msg)" -ForegroundColor Red
            continue
        }
        $emp = $resp.data.employees | Select-Object -First 1
        if (-not $emp) {
            Write-Host "  ⚠️  未找到工号 $JobNumber 的员工" -ForegroundColor Yellow
            continue
        }
        Write-Host "  返回的 base_info 字段:" -ForegroundColor Green
        if ($emp.base_info) {
            $emp.base_info.PSObject.Properties | ForEach-Object {
                $v = $_.Value
                $vs = if ($null -eq $v) { '<null>' } else { ($v | ConvertTo-Json -Depth 3 -Compress) }
                if ($vs.Length -gt 120) { $vs = $vs.Substring(0,120) + '...' }
                Write-Host "    $($_.Name) = $vs" -ForegroundColor Gray
            }
        } else { Write-Host "    <整个 base_info 为空>" -ForegroundColor DarkGray }

        Write-Host "  返回的 work_info 字段:" -ForegroundColor Green
        if ($emp.work_info) {
            $emp.work_info.PSObject.Properties | ForEach-Object {
                $v = $_.Value
                $vs = if ($null -eq $v) { '<null>' } else { ($v | ConvertTo-Json -Depth 3 -Compress) }
                if ($vs.Length -gt 120) { $vs = $vs.Substring(0,120) + '...' }
                Write-Host "    $($_.Name) = $vs" -ForegroundColor Gray
            }
        } else { Write-Host "    <整个 work_info 为空>" -ForegroundColor DarkGray }
    } catch {
        Write-Host "  ❌ 请求异常: $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "────────────────────────────────────────────────" -ForegroundColor Cyan
Write-Host "诊断说明:" -ForegroundColor Cyan
Write-Host "  - 如果 work_info.email 测试里 email 字段 <null>，说明 scope 没生效" -ForegroundColor Yellow
Write-Host "  - 如果某一行连字段本身都不出现（整个 PSObject.Properties 里找不到 email），说明飞书把它静默剔除了 = 权限问题" -ForegroundColor Yellow
Write-Host "  - 如果 base_info.mobile 也没返回，确认是权限范围问题，需重发 app 版本" -ForegroundColor Yellow

Write-Log "===== scope 诊断结束 =====" -Level OK

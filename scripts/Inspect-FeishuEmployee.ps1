# Inspect-FeishuEmployee.ps1 —— 诊断飞书 employee 对象真实字段结构
# 专门看 email 字段到底是字符串、对象还是数组
#
# 用法：.\scripts\Inspect-FeishuEmployee.ps1
#       .\scripts\Inspect-FeishuEmployee.ps1 -Count 5
#       .\scripts\Inspect-FeishuEmployee.ps1 -JobNumber <工号>

param(
    [int]$Count = 3,
    [string]$JobNumber,
    [string]$ConfigPath
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
. (Join-Path $root 'lib\Common.ps1')
. (Join-Path $root 'lib\Logger.ps1')
. (Join-Path $root 'lib\Feishu-Api.ps1')

Start-SyncLog -Tag 'inspect'
Write-SectionHeader "飞书员工字段诊断"

$config = Import-SyncConfig -ConfigPath $ConfigPath
$token = Get-FeishuTenantAccessToken -Config $config
$depts = Get-FeishuAllDepartments -Config $config -Token $token
$deptIds = @($depts | ForEach-Object { $_.OpenId })
$rawEmps = Get-FeishuEmployeesByDeptIds -Config $config -Token $token -DeptOpenIds $deptIds
Write-Log "拉到在职员工总数: $($rawEmps.Count)" -Level OK

# 过滤：指定工号 or 取前 N 个
if ($JobNumber) {
    $picks = @($rawEmps | Where-Object { [string]$_.work_info.job_number -eq $JobNumber })
    if ($picks.Count -eq 0) { Write-Log "飞书里没找到工号 $JobNumber" -Level WARN; return }
} else {
    $picks = @($rawEmps | Select-Object -First $Count)
}

function Show-FieldType {
    param($Label, $Value)
    if ($null -eq $Value) {
        Write-Host "  ${Label}: <null>" -ForegroundColor DarkGray
    } else {
        $t = $Value.GetType().FullName
        Write-Host "  ${Label}: type=$t" -ForegroundColor Yellow
        $Value | ConvertTo-Json -Depth 5 -Compress | Write-Host -ForegroundColor Gray
    }
}

foreach ($e in $picks) {
    $realName = $null
    if ($e.base_info.name.name.default_value) { $realName = [string]$e.base_info.name.name.default_value }
    elseif ($e.base_info.name.default_value)  { $realName = [string]$e.base_info.name.default_value }

    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  员工: $realName  工号: $([string]$e.work_info.job_number)" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan

    Write-Host ""
    Write-Host "[候选 email 字段]" -ForegroundColor Cyan
    Show-FieldType 'base_info.email'              $e.base_info.email
    Show-FieldType 'base_info.emails'             $e.base_info.emails
    Show-FieldType 'base_info.mobile'             $e.base_info.mobile
    Show-FieldType 'work_info.email'              $e.work_info.email
    Show-FieldType 'work_info.work_email'         $e.work_info.work_email
    Show-FieldType 'work_info.emails'             $e.work_info.emails
    Show-FieldType 'work_info.work_station.email' $e.work_info.work_station.email

    Write-Host ""
    Write-Host "[base_info 顶层字段名列表]" -ForegroundColor Cyan
    $e.base_info.PSObject.Properties.Name | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }

    Write-Host ""
    Write-Host "[work_info 顶层字段名列表]" -ForegroundColor Cyan
    $e.work_info.PSObject.Properties.Name | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }

    Write-Host ""
    Write-Host "[完整 raw JSON]" -ForegroundColor DarkGray
    $e | ConvertTo-Json -Depth 10
}

Write-Log "===== 诊断结束 =====" -Level OK

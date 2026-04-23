# Feishu → Active Directory Sync

[English](#english) | [中文](#中文)

---

## 中文

基于 PowerShell 的单向同步工具，将**飞书（Lark）**组织架构与在职员工同步到 **Microsoft Active Directory**。

> 单向同步：飞书是权威源，AD 是镜像。飞书改了 AD 跟着改，反向不处理。

### 一、功能特性

**组织架构同步**
- 从飞书拉取全量启用部门，按层级镜像到指定 AD 根 OU 下
- 部门改名 / 移动 / 新增 自动同步
- 部门稳定标识（`open_department_id`）写入 OU 的 `description` 字段，即使部门改名也能匹配回来
- 飞书删除部门时 AD OU 保留，不主动删（防误删 OU 下的用户）

**员工账号同步**
- 匹配键：飞书工号 ↔ AD `employeeID`
- 新员工自动创建 AD 账号（`sAMAccountName = 工号`、`UPN = 工号@<你的域>`、`mail = 飞书邮箱`）
- 存量员工同步姓名、邮箱、部门、所在 OU 位置
- 飞书离职员工自动禁用并移至归档 OU，不删除
- **绝不修改**现有账号的密码和启用状态

**存量账号回填**
- 提供 `Match-FeishuToAD.ps1` 脚本，对 AD 中 `employeeID` 为空的账号，按工号 / 邮箱 / 姓名三级匹配回填，避免首次同步时账号重复

**自动化**
- 提供一键安装 Windows 计划任务的脚本

### 二、运行环境

- Windows Server（加域成员机，不必是 DC）或 Windows 10/11 专业版
- PowerShell 5.1+ 或 PowerShell 7+
- RSAT ActiveDirectory PowerShell 模块
  ```powershell
  # Windows Server
  Install-WindowsFeature RSAT-AD-PowerShell

  # Windows 10/11
  Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
  ```
- 可访问 `https://open.feishu.cn`
- 运行账号具备在目标 OU 下创建 / 修改 / 移动 AD 对象的权限

### 三、快速开始

#### 1. 飞书侧准备

在飞书开放平台创建企业自建应用，开通以下权限并**发布新版本**（关键：scope 只有发版后才生效）：

| Scope | 说明 |
|---|---|
| `contact:department.base:readonly` | 读部门基础信息 |
| `directory:employee:list` | 员工列表 |
| `directory:employee.base.name.name:read` | 读员工姓名 |
| `directory:employee.base.department:read` | 读员工所在部门 |
| `directory:employee.base.email:read` | 读个人邮箱 |
| `directory:employee.work.job_number:read` | 读工号 |
| `directory:employee.work.email:read` | 读工作邮箱 |

记下 `App ID` 和 `App Secret`。

#### 2. AD 侧准备

在 AD 中创建两个 OU（名称自定义，写入配置即可）：

- 同步根 OU：飞书部门树镜像的父节点，如 `OU=FeishuSync,DC=example,DC=com`
- 归档 OU：离职员工移入的位置，如 `OU=FeishuArchive,DC=example,DC=com`

#### 3. 本地部署

```powershell
# 克隆仓库
git clone <this-repo-url> C:\FeishuToAD
cd C:\FeishuToAD

# 生成配置文件
Copy-Item config.sample.json config.json
notepad config.json

# 允许本地脚本执行
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

#### 4. 填写 `config.json`

```json
{
  "feishu": {
    "appId": "cli_xxxxxxxxxxxxxxxx",
    "appSecret": "<your-feishu-app-secret>",
    "apiBase": "https://open.feishu.cn"
  },
  "ad": {
    "domain": "EXAMPLE.COM",
    "upnSuffix": "@EXAMPLE.COM",
    "syncRootOu": "OU=FeishuSync,DC=example,DC=com",
    "archiveOu": "OU=FeishuArchive,DC=example,DC=com"
  },
  "user": {
    "defaultPassword": "ChangeMe!2024",
    "passwordNeverExpires": true,
    "changePasswordAtLogon": true,
    "enabledOnCreate": true
  },
  "dept": {
    "inactiveOuPrefix": "[Disabled]",
    "storeFeishuIdIn": "description"
  },
  "sync": {
    "logRetentionDays": 30
  }
}
```

字段说明：

| 字段 | 含义 |
|---|---|
| `feishu.appId` / `appSecret` | 飞书应用凭证 |
| `feishu.apiBase` | 飞书 OpenAPI 根地址，国际版请用 `https://open.larksuite.com` |
| `ad.domain` / `upnSuffix` | 你的 AD 域名，用于拼接 UPN |
| `ad.syncRootOu` | 同步根 OU 的完整 DN |
| `ad.archiveOu` | 离职员工归档 OU 的完整 DN |
| `user.defaultPassword` | 新建账号的初始密码，需满足域密码策略 |
| `user.passwordNeverExpires` | 新建账号密码是否永不过期 |
| `user.changePasswordAtLogon` | 新建账号首次登录是否强制改密 |
| `user.enabledOnCreate` | 新建账号是否启用 |
| `dept.inactiveOuPrefix` | 预留字段，当前未使用 |
| `dept.storeFeishuIdIn` | OU 存飞书 ID 用的字段，默认 `description` |
| `sync.logRetentionDays` | 日志保留天数 |

**加固建议**：配置文件含密钥，限制访问权限

```powershell
icacls config.json /inheritance:r /grant:r "SYSTEM:F" "Administrators:F"
```

#### 5. 按阶段验证

```powershell
# 阶段 1：预览飞书部门树（不连 AD）
.\scripts\Test-DeptPreview.ps1

# 阶段 2：审计 AD 现有账号 employeeID 填写情况
.\scripts\Audit-ADEmployeeId.ps1 -SearchBase "OU=FeishuSync,DC=example,DC=com" -OutputCsv ".\ad-audit.csv"

# 阶段 3（可选）：存量账号工号回填
.\scripts\Match-FeishuToAD.ps1 -WhatIf
.\scripts\Match-FeishuToAD.ps1 -ApplyP1 -ApplyP2

# 阶段 4：仅同步部门 OU
.\scripts\Sync-DeptsOnly.ps1 -WhatIf
.\scripts\Sync-DeptsOnly.ps1

# 阶段 5：单工号测试
.\scripts\Test-SingleUser.ps1 -EmployeeNo <员工工号> -WhatIf
.\scripts\Test-SingleUser.ps1 -EmployeeNo <员工工号>

# 阶段 6：完整同步
.\Sync-FeishuToAD.ps1 -Mode Full -WhatIf
.\Sync-FeishuToAD.ps1 -Mode Full

# 阶段 7：安装计划任务（管理员 PowerShell）
.\scripts\Install-ScheduledTask.ps1
```

### 四、运行模式

主脚本 `Sync-FeishuToAD.ps1` 通过 `-Mode` 参数选择：

| Mode | 行为 |
|---|---|
| `Preview` | 仅打印飞书部门树，不连 AD |
| `DeptsOnly` | 同步部门 OU 树，不动用户 |
| `SingleUser` | 仅同步一个指定工号，用于测试字段映射 |
| `Full` | 完整同步：部门 + 全量员工 + 离职差集归档 |

所有模式都支持 `-WhatIf` 干跑，日志里 `[DRY]` 前缀标出"将要做但没做"的动作。

### 五、字段映射

#### 部门

| 飞书来源 | AD 目标 |
|---|---|
| `open_department_id` | OU 的 `description = feishu:<openid>` |
| `name.default_value` | OU 的 `Name` |
| `parent_department_id` | OU 在同步树中的层级位置 |

#### 员工

| 飞书来源 | AD 目标 |
|---|---|
| `work_info.job_number` | `sAMAccountName`、`employeeID` |
| `work_info.job_number + upnSuffix` | `UserPrincipalName` |
| `base_info.name.name.default_value` | `cn`、`displayName`、`sn` |
| `base_info.email`（主）/ `work_info.email`（备） | `mail` |
| `base_info.departments[0].department_id` | 所在 OU 位置 + 该部门名写入 `department` |

### 六、行为边界

| 场景 | 行为 |
|---|---|
| 飞书新员工，AD 无此工号 | 新建 AD 账号 |
| 飞书员工改名 | 更新 `displayName` / `sn` / `cn` |
| 飞书员工换部门 | 更新 `department` + `Move-ADObject` |
| 飞书员工改邮箱 | 更新 `mail`（UPN 保持不变） |
| 飞书员工离职 | `Disable-ADAccount` + 移入归档 OU |
| 飞书员工重新入职 | 更新属性 + 挪回部门 OU；**不自动启用**，需人工审核 |
| AD 手工账号（无 `employeeID`） | 完全不动 |
| AD 有 `employeeID` 但在 `syncRootOu` 之外 | 会被移入同步树（注意） |
| 飞书部门删除 | AD OU 保留不删 |

### 七、安全原则

代码层面已保证：

- 已有账号的**密码**绝不修改（Update 分支内无 `-AccountPassword`）
- 已有账号的**启用状态**绝不修改（Update 分支内无 `Enable-ADAccount`/`Disable-ADAccount`）
- 账号匹配基于 `employeeID` 精确等值，不做姓名模糊匹配（`Match-FeishuToAD.ps1` 的 P3 姓名匹配默认不应用）
- 离职差集扫描只限于 `syncRootOu` 范围内，范围外账号永远不会被脚本禁用

### 八、日志

- 位置：`logs\yyyyMMdd-<tag>.log`（UTF-8）
- Tag：`dept-preview` / `depts-only` / `user-<工号>` / `full-sync` / `audit` / `match` / `scope` / `inspect`
- 控制台同步输出
- 保留天数由 `config.json` 的 `sync.logRetentionDays` 控制

级别约定：

```
[INFO]  普通信息
[OK]    成功操作
[WARN]  警告，不中断
[ERR]   错误
[DRY]   WhatIf 模式下"将要做但没做"
```

### 九、目录结构

```
feishutoad/
├── README.md                           本文档
├── LICENSE
├── config.sample.json                  配置模板
├── config.json                         真实配置（gitignore）
├── .gitignore
├── Sync-FeishuToAD.ps1                 主脚本入口
│
├── lib/
│   ├── Common.ps1                      编码与配置加载
│   ├── Logger.ps1                      分级日志
│   ├── Feishu-Api.ps1                  飞书 API 封装
│   └── AD-Operations.ps1               AD 增删改封装
│
├── scripts/
│   ├── Test-DeptPreview.ps1            预览飞书部门树
│   ├── Sync-DeptsOnly.ps1              仅同步部门 OU
│   ├── Test-SingleUser.ps1             单工号测试
│   ├── Audit-ADEmployeeId.ps1          AD employeeID 填写情况审计
│   ├── Match-FeishuToAD.ps1            存量账号 employeeID 回填
│   ├── Install-ScheduledTask.ps1       安装/卸载计划任务
│   ├── Inspect-FeishuEmployee.ps1      员工字段结构诊断
│   └── Inspect-FeishuScope.ps1         飞书 scope 生效情况诊断
│
└── logs/                               运行日志（gitignore）
```

### 十、常见问题

| 现象 | 原因 / 处理 |
|---|---|
| `未安装 ActiveDirectory PowerShell 模块` | 安装 RSAT（见第二节） |
| `同步根 OU 不存在` | 检查 `config.json` 的 `syncRootOu` DN 与 AD 实际是否一致 |
| `The password does not meet the password policy requirements` | 域密码策略过严，调整 `defaultPassword` 或放宽策略 |
| `2220009 Filter field is invalid` | 飞书 scope 未开通或开通后未发布版本。用 `Inspect-FeishuScope.ps1` 定位 |
| `Insufficient access rights` | 运行账号对目标 OU 无权限，改用有权账号或做 AD 委派 |
| 邮箱字段返回 null | 开通 `directory:employee.base.email:read` 后必须发布新版本 scope 才生效 |
| 姓名显示为 `@{default_value=...}` | 应使用 `base_info.name.name.default_value` 路径取值 |
| 中文乱码 | 改用 PowerShell 7，或在脚本开头加 `chcp 65001`；`.ps1` 文件保存为 UTF-8 BOM |

### 十一、计划任务

```powershell
# 安装（管理员 PowerShell）
.\scripts\Install-ScheduledTask.ps1

# 手动触发
Start-ScheduledTask -TaskName 'FeishuToAD-Sync'

# 查看状态
Get-ScheduledTaskInfo -TaskName 'FeishuToAD-Sync'

# 卸载
.\scripts\Install-ScheduledTask.ps1 -Remove
```

默认：每小时整点执行 `Sync-FeishuToAD.ps1 -Mode Full`，最长执行 30 分钟。

### 十二、贡献

欢迎 Issue 和 Pull Request。提交代码前请确保：

- 不要把真实的 `config.json`、`appSecret`、域名、OU DN 写进 commit
- 日志样例中替换真实员工信息
- 新增 scope 依赖要更新 README 的权限列表

### 十三、License

MIT License —— 见 [LICENSE](./LICENSE) 文件。

---

## English

A PowerShell-based one-way sync tool that syncs **Feishu (Lark)** organizational structure and active employees to **Microsoft Active Directory**.

> One-way sync: Feishu is the authoritative source, AD is a mirror. Changes in Feishu propagate to AD; the reverse is not handled.

### Features

- **Department sync**: Full department tree mirrored under a configurable root OU, supports rename / move / add; stable Feishu `open_department_id` stored in OU `description` attribute for reliable re-matching after renames
- **User sync**: Matched by Feishu job number ↔ AD `employeeID`; creates new accounts, updates existing (name / email / department / OU location), disables and archives departed employees
- **Safety**: Never touches passwords or account-enabled state of existing users; match strictly by `employeeID`
- **Backfill**: `Match-FeishuToAD.ps1` for reconciling existing AD accounts that have empty `employeeID` before first sync
- **Automation**: One-click scheduled task installation

### Requirements

- Windows Server (domain-joined, not necessarily DC) or Windows 10/11 Pro
- PowerShell 5.1+ or 7+
- RSAT ActiveDirectory module
- Network access to `https://open.feishu.cn` (or `https://open.larksuite.com` for international edition)
- Account with permissions to create/modify/move AD objects in target OUs

### Quick Start

1. Create a Feishu custom app with scopes listed in the Chinese section above, **publish a new version** (scopes only take effect after version publish)
2. Create two OUs in AD: a sync root and an archive OU
3. Clone repo, copy `config.sample.json` to `config.json`, fill in your values
4. Run staged validation: `Test-DeptPreview.ps1` → `Audit-ADEmployeeId.ps1` → `Sync-DeptsOnly.ps1 -WhatIf` → `Test-SingleUser.ps1` → `Sync-FeishuToAD.ps1 -Mode Full -WhatIf` → full run
5. Install scheduled task via `Install-ScheduledTask.ps1`

### Modes

`Sync-FeishuToAD.ps1 -Mode <Preview|DeptsOnly|SingleUser|Full> [-EmployeeNo <id>] [-WhatIf]`

### Security Note

Do not commit `config.json` or real credentials. The `.gitignore` already excludes `config.json`, logs, and secrets.

### License

MIT

<# 
主入口：权限检查与模块加载 
#>

# 检查管理员权限
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`"" -Verb RunAs
    exit
}

# 加载依赖组件（必须先加载系统组件）
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# 定义全局共享变量（所有模块可访问）
$script:domainContext = $null          # 域控连接上下文
$script:remoteSession = $null
$script:currentOU = $null
$script:allUsersOU = $null
$script:allUsers = New-Object System.Collections.ArrayList  # 所有用户数据（原始未过滤）
$script:filteredUsers = New-Object System.Collections.ArrayList  # 过滤后用户数据（用于分页）
$script:allGroups = New-Object System.Collections.ArrayList # 所有组数据（原始未过滤）
$script:filteredGroups = New-Object System.Collections.ArrayList # 过滤后组数据（用于分页）
$script:originalGroupSamAccount = $null # 原始组账号（用于修改组）
$script:connectionStatus = "未连接到域控" # 连接状态
$script:userCountStatus = "0"           # 用户计数
$script:groupCountStatus = "0"          # 组计数

# ---------------------- 分页功能全局变量 ----------------------
$script:pageSize = 6                  # 每页显示条数（固定6条）
# 用户列表分页状态
$script:currentUserPage = 1             # 当前用户页码
$script:totalUserPages = 1              # 用户总页数
# 组列表分页状态
$script:currentGroupPage = 1            # 当前组页码
$script:totalGroupPages = 1             # 组总页数

$script:defaultShowAll = $true  # 控制用户列表默认全显
$script:groupDefaultShowAll = $true  # 控制组列表默认全显

# 加载工具类（Utilities）
. "$PSScriptRoot/Utilities/Helpers.ps1"
. "$PSScriptRoot/Utilities/PinyinConverter.ps1"
. "$PSScriptRoot/Utilities/ImportExportUsers.ps1"

# 加载核心函数（Functions）
. "$PSScriptRoot/Functions/DomainOperations.ps1"
. "$PSScriptRoot/Functions/UserOperations.ps1"
. "$PSScriptRoot/Functions/GroupOperations.ps1"
. "$PSScriptRoot/Functions/OUOperations.ps1"

# 加载窗体控件（Forms）
. "$PSScriptRoot/Forms/Controls/ConnectionPanel.ps1"
. "$PSScriptRoot/Forms/Controls/UserManagementPanel.ps1"
. "$PSScriptRoot/Forms/Controls/GroupManagementPanel.ps1"
. "$PSScriptRoot/Forms/Controls/StatusBar.ps1"
. "$PSScriptRoot/Forms/MainForm.ps1"

# 启动主窗口
$script:mainForm.ShowDialog() | Out-Null
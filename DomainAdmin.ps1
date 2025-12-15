<# 
主入口：权限检查与模块加载 
#>

# 0. 全局输出抑制（消除数字弹出）
$ErrorActionPreference = 'SilentlyContinue'   # 先抑制所有非关键输出
$InformationPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'

# 1. 强制STA线程（控制台模式关键，抑制返回值）
if (-not ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA')) {
    $null = [System.Threading.Thread]::CurrentThread.SetApartmentState('STA')
}

#2. 重定向控制台输出
if ($MyInvocation.MyCommand.CommandType -eq 'Application' -and -not $noConsole) {
    $null = [Console]::SetOut((New-Object System.IO.StreamWriter([System.IO.Stream]::Null)))
    $ErrorActionPreference = 'Continue'  # 仅保留关键错误
}

# 3. 适配PS1/EXE的路径处理（抑制所有返回值）
$script:AppRoot = $null
try {
    $null = $exePath = [System.Reflection.Assembly]::GetEntryAssembly().Location
    if ($exePath) {
        $null = $script:AppRoot = [System.IO.Path]::GetDirectoryName($exePath)
    }
} catch {
    if ($MyInvocation.MyCommand.CommandType -eq 'Application') {
        $null = $script:AppRoot = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
    } else {
        $null = $script:AppRoot = $PSScriptRoot
    }
}
if (-not $script:AppRoot -or $script:AppRoot -eq '') {
    $null = $script:AppRoot = [System.IO.Directory]::GetCurrentDirectory()
}

# 4. 管理员权限检查（隐藏新窗口）
if ($MyInvocation.MyCommand.CommandType -ne 'Application') {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`"" -Verb RunAs -WindowStyle Hidden
        exit
    }
}

#5. 强制提前加载Forms组件（完全抑制返回值）
try {
    # 加载核心程序集（无变量赋值，直接抑制返回值）
    $null = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $null = [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $null = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

    # 显式引用MessageBox类型（避免控制台模式下找不到）
    $null = [System.Windows.Forms.MessageBox]
    $ErrorActionPreference = 'Continue'  # 恢复错误输出，便于捕获关键异常
} catch {
    # 控制台模式下优先输出到控制台，再弹框
    Write-Error "组件加载失败：$_"
    if ([System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")) {
        [System.Windows.Forms.MessageBox]::Show("组件加载失败：$_", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    exit 1
}

# 6. 全局变量
$script:domainContext = $null          
$script:remoteSession = $null
$script:currentOU = $null
$script:allUsersOU = $null
$null = $script:allUsers = [System.Collections.ArrayList]@()          # 抑制初始化输出
$null = $script:filteredUsers = [System.Collections.ArrayList]@()
$null = $script:allGroups = [System.Collections.ArrayList]@()
$null = $script:filteredGroups = [System.Collections.ArrayList]@()
$script:originalGroupSamAccount = $null
$script:connectionStatus = "未连接到域控"
$script:userCountStatus = "0"           
$script:groupCountStatus = "0"          

$script:dynamicUserPageSize = 0
$script:dynamicGroupPageSize = 0
$script:currentUserPage = 1              
$script:totalUserPages = 1               
$script:currentGroupPage = 1             
$script:totalGroupPages = 1             

#7. 加载子脚本
$subScripts = @(
    (Join-Path -Path $script:AppRoot -ChildPath "Utilities\Helpers.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Utilities\PinyinConverter.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Utilities\ImportExportUsers.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Functions\DomainOperations.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Functions\UserOperations.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Functions\GroupOperations.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Functions\OUOperations.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Functions\RestrictLogin.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Functions\RestrictLogonTime.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Forms\Controls\ConnectionPanel.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Forms\Controls\UserManagementPanel.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Forms\Controls\GroupManagementPanel.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Forms\Controls\StatusBar.ps1"),
    (Join-Path -Path $script:AppRoot -ChildPath "Forms\MainForm.ps1")
)

foreach ($scriptPath in $subScripts) {
    if (-not (Test-Path -Path $scriptPath -PathType Leaf)) {
        $errorMsg = "子脚本文件不存在：`n$scriptPath`n`n请确保EXE与Utilities/Functions/Forms文件夹同目录！"
        Write-Error $errorMsg
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "文件缺失", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }
    try {
        $null = . $scriptPath  # 抑制子脚本加载的隐式输出
    } catch {
        $errorMsg = "加载子脚本失败：`n文件路径：$scriptPath`n错误信息：$_"
        Write-Error $errorMsg
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "加载失败", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }
}

#8. 启动主窗口
try {
    $null = $script:mainForm.ShowDialog()  # 彻底抑制输出
} catch {
    Write-Error "主窗口启动失败：$_"
    [System.Windows.Forms.MessageBox]::Show("主窗口启动失败：$_", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
} finally {
    [System.Windows.Forms.Application]::Exit()
    exit 0
}
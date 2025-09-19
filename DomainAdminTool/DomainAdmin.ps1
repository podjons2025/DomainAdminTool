<# 
����ڣ�Ȩ�޼����ģ����� 
#>

# ������ԱȨ��
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`"" -Verb RunAs
    exit
}

# ������������������ȼ���ϵͳ�����
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# ����ȫ�ֹ������������ģ��ɷ��ʣ�
$script:domainContext = $null          # �������������
$script:remoteSession = $null
$script:currentOU = $null
$script:allUsersOU = $null
$script:allUsers = New-Object System.Collections.ArrayList  # �����û�����
$script:allGroups = New-Object System.Collections.ArrayList # ����������
$script:originalGroupSamAccount = $null # ԭʼ���˺ţ������޸��飩
$script:connectionStatus = "δ���ӵ����" # ����״̬
$script:userCountStatus = "0"           # �û�����
$script:groupCountStatus = "0"          # �����

# ���ع����ࣨUtilities��
. "$PSScriptRoot/Utilities/Helpers.ps1"
. "$PSScriptRoot/Utilities/PinyinConverter.ps1"
. "$PSScriptRoot/Utilities/importExportUsers.ps1"

# ���غ��ĺ�����Functions��
. "$PSScriptRoot/Functions/DomainOperations.ps1"
. "$PSScriptRoot/Functions/UserOperations.ps1"
. "$PSScriptRoot/Functions/GroupOperations.ps1"
. "$PSScriptRoot/Functions/OUOperations.ps1"

# ���ش���ؼ���Forms��
. "$PSScriptRoot/Forms/Controls/ConnectionPanel.ps1"
. "$PSScriptRoot/Forms/Controls/UserManagementPanel.ps1"
. "$PSScriptRoot/Forms/Controls/GroupManagementPanel.ps1"
. "$PSScriptRoot/Forms/Controls/StatusBar.ps1"
. "$PSScriptRoot/Forms/MainForm.ps1"

# ����������
$script:mainForm.ShowDialog() | Out-Null
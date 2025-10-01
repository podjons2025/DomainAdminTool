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
$script:allUsers = New-Object System.Collections.ArrayList  # �����û����ݣ�ԭʼδ���ˣ�
$script:filteredUsers = New-Object System.Collections.ArrayList  # ���˺��û����ݣ����ڷ�ҳ��
$script:allGroups = New-Object System.Collections.ArrayList # ���������ݣ�ԭʼδ���ˣ�
$script:filteredGroups = New-Object System.Collections.ArrayList # ���˺������ݣ����ڷ�ҳ��
$script:originalGroupSamAccount = $null # ԭʼ���˺ţ������޸��飩
$script:connectionStatus = "δ���ӵ����" # ����״̬
$script:userCountStatus = "0"           # �û�����
$script:groupCountStatus = "0"          # �����

# ---------------------- ��ҳ����ȫ�ֱ��� ----------------------
$script:pageSize = 6                  # ÿҳ��ʾ�������̶�6����
# �û��б��ҳ״̬
$script:currentUserPage = 1             # ��ǰ�û�ҳ��
$script:totalUserPages = 1              # �û���ҳ��
# ���б��ҳ״̬
$script:currentGroupPage = 1            # ��ǰ��ҳ��
$script:totalGroupPages = 1             # ����ҳ��

$script:defaultShowAll = $true  # �����û��б�Ĭ��ȫ��
$script:groupDefaultShowAll = $true  # �������б�Ĭ��ȫ��

# ���ع����ࣨUtilities��
. "$PSScriptRoot/Utilities/Helpers.ps1"
. "$PSScriptRoot/Utilities/PinyinConverter.ps1"
. "$PSScriptRoot/Utilities/ImportExportUsers.ps1"

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
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
#$script:fixedPageSize = 6                 # �̶�Ĭ�Ϸ�ҳ��С
# ��̬��ҳ��С���洰�ڿ�������仯�����ı�����
$script:dynamicUserPageSize = 0  # �û��б�̬����
$script:dynamicGroupPageSize = 0 # ���б�̬����

# ����ԭ�б��������䣩
$script:currentUserPage = 1              # ��ǰ�û�ҳ��
$script:totalUserPages = 1               # �û���ҳ��
$script:currentGroupPage = 1             # ��ǰ��ҳ��
$script:totalGroupPages = 1              # ����ҳ��


# ���ع����ࣨUtilities��
. "$PSScriptRoot/Utilities/Helpers.ps1"
. "$PSScriptRoot/Utilities/PinyinConverter.ps1"
. "$PSScriptRoot/Utilities/ImportExportUsers.ps1"


# ���غ��ĺ�����Functions��
. "$PSScriptRoot/Functions/DomainOperations.ps1"
. "$PSScriptRoot/Functions/UserOperations.ps1"
. "$PSScriptRoot/Functions/GroupOperations.ps1"
. "$PSScriptRoot/Functions/OUOperations.ps1"
. "$PSScriptRoot/Functions/RestrictLogin.ps1"
. "$PSScriptRoot/Functions/RestrictLogonTime.ps1"


# ���ش���ؼ���Forms��
. "$PSScriptRoot/Forms/Controls/ConnectionPanel.ps1"
. "$PSScriptRoot/Forms/Controls/UserManagementPanel.ps1"
. "$PSScriptRoot/Forms/Controls/GroupManagementPanel.ps1"
. "$PSScriptRoot/Forms/Controls/StatusBar.ps1"
. "$PSScriptRoot/Forms/MainForm.ps1"

<#
.SYNOPSIS
ɾ����ǰ�ű�����Ŀ¼���������ݣ���������
#>
function deleteapp {
    # ��ȡ��ǰ�ű�����Ŀ¼������·��������·���е������ַ��������������
    $targetDir = [System.IO.Path]::GetFullPath($PSScriptRoot)

    # ������̨PowerShell���̣��ӳ�2���ɾ�����Ƴ������exit
    Start-Process -FilePath powershell.exe -ArgumentList @(
        "-NoProfile", "-ExecutionPolicy Bypass",  # ���ٻ�������
        "-Command", "Start-Sleep -Seconds 2; Remove-Item -Path '$targetDir' -Recurse -Force -ErrorAction SilentlyContinue"
    ) -NoNewWindow -PassThru | Out-Null

}

$script:mainForm.Add_Closed({
    # ������ȫ�رա���Դ�ͷź���ִ��ɾ���߼�
    deleteapp
})


# ����������
$script:mainForm.ShowDialog() | Out-Null

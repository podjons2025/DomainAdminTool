<# 
�޸����Ӻ��������OU��ع��ܼ��� 
#>

function ConnectToDomain {
    $selectedDomain = $script:comboDomain.SelectedItem
    if (-not $selectedDomain) {
        [System.Windows.Forms.MessageBox]::Show("��ѡ��һ����ص�ַ", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
	
    $domain = $selectedDomain.Server
    $adminUser = $script:textAdmin.Text
    $adminPassword = $script:textPassword.Text

    if ([string]::IsNullOrEmpty($adminUser) -or [string]::IsNullOrEmpty($adminPassword)) {
        [System.Windows.Forms.MessageBox]::Show("����д����Ա�˺ź�����", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
	
    try {
		# ����ƾ��
        $securePassword = ConvertTo-SecureString $adminPassword -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential ($adminUser, $securePassword)
        
        # ����Զ�̻Ự
        $script:remoteSession = New-PSSession -ComputerName $domain -Credential $credential -ErrorAction Stop
     				
        $script:connectionStatus = "������֤Զ�����AD����..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # Զ����֤ADģ��
        $domainInfo = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            Import-Module ActiveDirectory -ErrorAction Stop
            return Get-ADDomain -ErrorAction Stop
        } -ErrorAction Stop
        
        $script:domainContext = @{
            Server = $domain
            Credential = $credential
            DomainInfo = $domainInfo
        }
        
        # ����Ĭ��OU
        $script:currentOU = "CN=Users,$($domainInfo)"
        $script:textOU.Text = $script:currentOU
		
        
        $script:connectionStatus = "�����ӵ����: $($selectedDomain.Name)��Զ��ִ�У�"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        
        # ���°�ť״̬
        $script:buttonConnect.Enabled = $false
        $script:buttonConnect.BackColor = [System.Drawing.Color]::FromArgb(169, 169, 169)
        $script:buttonDisconnect.Enabled = $true
        $script:buttonDisconnect.BackColor = [System.Drawing.Color]::FromArgb(220, 20, 60)        
        
        $script:comboDomain.Enabled = $false
        $script:textAdmin.Enabled = $false
        $script:textPassword.Enabled = $false
        
        # ��������
        LoadUserList   # ����UserOperations.ps1
        LoadGroupList  # ����GroupOperations.ps1
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "WinRM") { $errorMsg += "`n��ȷ��Զ�����������WinRM��������winrm quickconfig��" }
        elseif ($errorMsg -match "ActiveDirectory") { $errorMsg += "`n��ȷ��Զ������Ѱ�װADģ��" }
        $script:connectionStatus = "����ʧ��: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        $script:domainContext = $null
    }
}

function DisconnectFromDomain {
    $result = [System.Windows.Forms.MessageBox]::Show("ȷ��Ҫ�Ͽ�����ص�������", "ȷ��", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    
    $script:domainContext = $null
    $script:allUsers.Clear()
    $script:userDataGridView.DataSource = $null
    $script:allGroups.Clear()
    $script:groupDataGridView.DataSource = $null
    $script:allOUs = $null  # ���OU�б�
    
    $script:connectionStatus = "δ���ӵ����"
    $script:userCountStatus = "0"
    $script:groupCountStatus = "0"
    UpdateStatusBar
    $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::Black
    
    # ���°�ť״̬
    $script:buttonConnect.Enabled = $true
    $script:buttonConnect.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $script:buttonDisconnect.Enabled = $false
    $script:buttonDisconnect.BackColor = [System.Drawing.Color]::FromArgb(169, 169, 169)    
    
    $script:comboDomain.Enabled = $true
    $script:textAdmin.Enabled = $true
    $script:textPassword.Enabled = $true
    $script:textOU.Text = ""
    
    ClearInputFields
    ClearGroupInputFields
    $script:originalGroupSamAccount = $null
}

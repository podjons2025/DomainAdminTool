<# 
修改连接函数，添加OU相关功能激活 
#>

function ConnectToDomain {
    $selectedDomain = $script:comboDomain.SelectedItem
    if (-not $selectedDomain) {
        [System.Windows.Forms.MessageBox]::Show("请选择一个域控地址", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
	
    $domain = $selectedDomain.Server
    $adminUser = $script:textAdmin.Text
    $adminPassword = $script:textPassword.Text

    if ([string]::IsNullOrEmpty($adminUser) -or [string]::IsNullOrEmpty($adminPassword)) {
        [System.Windows.Forms.MessageBox]::Show("请填写管理员账号和密码", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
	
    try {
		# 创建凭据
        $securePassword = ConvertTo-SecureString $adminPassword -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential ($adminUser, $securePassword)
        
        # 建立远程会话
        $script:remoteSession = New-PSSession -ComputerName $domain -Credential $credential -ErrorAction Stop
     				
        $script:connectionStatus = "正在验证远程域控AD功能..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 远程验证AD模块
        $domainInfo = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            Import-Module ActiveDirectory -ErrorAction Stop
            return Get-ADDomain -ErrorAction Stop
        } -ErrorAction Stop
        
        $script:domainContext = @{
            Server = $domain
            Credential = $credential
            DomainInfo = $domainInfo
        }
        
        # 设置默认OU
        $script:currentOU = "CN=Users,$($domainInfo)"
        $script:textOU.Text = $script:currentOU
		
        
        $script:connectionStatus = "已连接到域控: $($selectedDomain.Name)（远程执行）"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        
        # 更新按钮状态
        $script:buttonConnect.Enabled = $false
        $script:buttonConnect.BackColor = [System.Drawing.Color]::FromArgb(169, 169, 169)
        $script:buttonDisconnect.Enabled = $true
        $script:buttonDisconnect.BackColor = [System.Drawing.Color]::FromArgb(220, 20, 60)        
        
        $script:comboDomain.Enabled = $false
        $script:textAdmin.Enabled = $false
        $script:textPassword.Enabled = $false
        
        # 加载数据
        LoadUserList   # 来自UserOperations.ps1
        LoadGroupList  # 来自GroupOperations.ps1
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "WinRM") { $errorMsg += "`n请确保远程域控已启用WinRM服务（运行winrm quickconfig）" }
        elseif ($errorMsg -match "ActiveDirectory") { $errorMsg += "`n请确保远程域控已安装AD模块" }
        $script:connectionStatus = "连接失败: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        $script:domainContext = $null
    }
}

function DisconnectFromDomain {
    $result = [System.Windows.Forms.MessageBox]::Show("确定要断开与域控的连接吗？", "确认", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    
    $script:domainContext = $null
    $script:allUsers.Clear()
    $script:userDataGridView.DataSource = $null
    $script:allGroups.Clear()
    $script:groupDataGridView.DataSource = $null
    $script:allOUs = $null  # 清除OU列表
    
    $script:connectionStatus = "未连接到域控"
    $script:userCountStatus = "0"
    $script:groupCountStatus = "0"
    UpdateStatusBar
    $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::Black
    
    # 更新按钮状态
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

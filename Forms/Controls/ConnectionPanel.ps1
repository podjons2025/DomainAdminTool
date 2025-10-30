<# 
域控连接设置面板 
#>

$script:connectionPanel = New-Object System.Windows.Forms.GroupBox
$script:connectionPanel.Text = "域控连接设置"
$script:connectionPanel.Dock = "Fill"
$script:connectionPanel.Padding = 5

# 连接区表格布局
$script:connectionTable = New-Object System.Windows.Forms.TableLayoutPanel
$script:connectionTable.Dock = "Fill"
$script:connectionTable.RowCount = 4 
$script:connectionTable.ColumnCount = 4
$script:connectionTable.Padding = 5
# 调整行比例
$script:connectionTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
$script:connectionTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
$script:connectionTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
$script:connectionTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
# 列布局
$script:connectionTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100)))
$script:connectionTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$script:connectionTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100)))
$script:connectionTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

# 1. 域控地址下拉框
$script:labelDomain = New-Object System.Windows.Forms.Label
$script:labelDomain.Text = "域控地址:"
$script:labelDomain.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:connectionTable.Controls.Add($script:labelDomain, 0, 0)

$script:comboDomain = New-Object System.Windows.Forms.ComboBox
$script:comboDomain.Dock = "Fill"
$script:comboDomain.DropDownStyle = "DropDownList"
$script:comboDomain.DisplayMember = "Name"
$script:comboDomain.ValueMember = "Server"
$script:comboDomain.Items.AddRange(@(	
    [PSCustomObject]@{Name = "域控（广州）- serverAD.abc.com"; Server = "serverAD.abc.com"; SystemAccount= "abc\admin"; Password = "Abc123456"},
    [PSCustomObject]@{Name = "域控（上海）- abc03.abc01.com"; Server = "abc03.abc01.com"; SystemAccount= "abc01\administrator"; Password = "Password123"},
    [PSCustomObject]@{Name = "测试域控（北京）- serverAD3.abc03.com"; Server = "serverAD3.abc03.com"; SystemAccount= "abc03\admin"; Password = ""}		
))
$script:comboDomain.SelectedIndex = 0
$script:comboDomain.Add_SelectedIndexChanged({
    $selectedDomain = $script:comboDomain.SelectedItem
    if ($selectedDomain) {
        $script:textAdmin.Text = $selectedDomain.SystemAccount 
        $script:textPassword.Text = $selectedDomain.Password
    }
})
$script:connectionTable.Controls.Add($script:comboDomain, 1, 0)

# 2. 管理员账号
$script:labelAdmin = New-Object System.Windows.Forms.Label
$script:labelAdmin.Text = "管理员账号:"
$script:labelAdmin.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:connectionTable.Controls.Add($script:labelAdmin, 2, 0)

$script:textAdmin = New-Object System.Windows.Forms.TextBox
$script:textAdmin.Text = $script:comboDomain.SelectedItem.SystemAccount 
$script:textAdmin.Dock = "Fill"
$script:connectionTable.Controls.Add($script:textAdmin, 3, 0)

# 3. 密码输入框
$script:labelPassword = New-Object System.Windows.Forms.Label
$script:labelPassword.Text = "密码:"
$script:labelPassword.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:connectionTable.Controls.Add($script:labelPassword, 0, 1)

$script:textPassword = New-Object System.Windows.Forms.TextBox
$script:textPassword.PasswordChar = '*'
$script:textPassword.Text = $script:comboDomain.SelectedItem.Password
$script:textPassword.Dock = "Fill"

# 添加鼠标按下事件 - 显示密码
$script:textPassword.Add_MouseDown({
    $this.PasswordChar = $null  # 取消密码掩码，显示明文
})

# 添加鼠标释放事件 - 恢复掩码
$script:textPassword.Add_MouseUp({
    $this.PasswordChar = '*'    # 恢复密码掩码
})

# 添加鼠标离开事件 - 确保离开时恢复掩码
$script:textPassword.Add_MouseLeave({
    $this.PasswordChar = '*'    # 恢复密码掩码
})

$script:connectionTable.Controls.Add($script:textPassword, 1, 1)

# 4. OU组织
$script:labelOU = New-Object System.Windows.Forms.Label
$script:labelOU.Text = "OU组织:"
$script:labelOU.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:connectionTable.Controls.Add($script:labelOU, 0, 2)  # 第3行第1列

$script:textOU = New-Object System.Windows.Forms.TextBox
$script:textOU.Dock = "Fill"
$script:textOU.ReadOnly = $true
$script:connectionTable.Controls.Add($script:textOU, 1, 2)  # 第3行第2列

# 5. 连接/断开按钮面板
$script:buttonPanel = New-Object System.Windows.Forms.Panel
$script:buttonPanel.Dock = "Fill"
$script:buttonPanel.Padding = 5

$script:buttonConnect = New-Object System.Windows.Forms.Button
$script:buttonConnect.Text = "连接域控"
$script:buttonConnect.Location = New-Object System.Drawing.Point(5, 5)
$script:buttonConnect.Width = 80
$script:buttonConnect.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
$script:buttonConnect.ForeColor = [System.Drawing.Color]::White
$script:buttonConnect.FlatStyle = "Flat"
$script:buttonConnect.Add_Click({ ConnectToDomain })  # 来自DomainOperations.ps1

$script:buttonDisconnect = New-Object System.Windows.Forms.Button
$script:buttonDisconnect.Text = "退出连接"
$script:buttonDisconnect.Location = New-Object System.Drawing.Point(115, 5)
$script:buttonDisconnect.Width = 80
$script:buttonDisconnect.BackColor = [System.Drawing.Color]::FromArgb(169, 169, 169)
$script:buttonDisconnect.ForeColor = [System.Drawing.Color]::White
$script:buttonDisconnect.FlatStyle = "Flat"
$script:buttonDisconnect.Enabled = $false
$script:buttonDisconnect.Add_Click({ DisconnectFromDomain })  # 来自DomainOperations.ps1

$script:buttonPanel.Controls.Add($script:buttonConnect)
$script:buttonPanel.Controls.Add($script:buttonDisconnect)
$script:connectionTable.Controls.Add($script:buttonPanel, 3, 1)  # 保持在密码行的右侧




# 6. OU操作按钮面板
$script:ouButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$script:ouButtonPanel.Dock = "Fill"  # 填充主面板的整行
$script:ouButtonPanel.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)  # 左右10px、上下5px内边距，优化视觉
$script:ouButtonPanel.FlowDirection = "LeftToRight"  # 按钮从左到右排列
$script:ouButtonPanel.WrapContents = $false  # 禁止按钮换行，保持一行显示
$script:ouButtonPanel.AutoScroll = $false  # 无需滚动条
$script:ouButtonPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)  # 与主窗体背景色一致，视觉统一

$script:buttonSwitchOU = New-Object System.Windows.Forms.Button
$script:buttonSwitchOU.Text = "切换OU组织"
$script:buttonSwitchOU.Width = 100
$script:buttonSwitchOU.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
$script:buttonSwitchOU.ForeColor = [System.Drawing.Color]::White
$script:buttonSwitchOU.FlatStyle = "Flat"
$script:buttonSwitchOU.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)  # 按钮间距
$script:buttonSwitchOU.Add_Click({ SwitchOU })

$script:buttonCreateOU = New-Object System.Windows.Forms.Button
$script:buttonCreateOU.Text = "新建OU组织"
$script:buttonCreateOU.Width = 100
$script:buttonCreateOU.BackColor = [System.Drawing.Color]::FromArgb(128, 0, 128)
$script:buttonCreateOU.ForeColor = [System.Drawing.Color]::White
$script:buttonCreateOU.FlatStyle = "Flat"
$script:buttonCreateOU.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:buttonCreateOU.Add_Click({ CreateNewOU })

$script:buttonRenameOU = New-Object System.Windows.Forms.Button
$script:buttonRenameOU.Text = "重命名OU组织"
$script:buttonRenameOU.Width = 100
$script:buttonRenameOU.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)  # 钢蓝色
$script:buttonRenameOU.ForeColor = [System.Drawing.Color]::White
$script:buttonRenameOU.FlatStyle = "Flat"
$script:buttonRenameOU.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:buttonRenameOU.Add_Click({ RenameExistingOU })

$script:buttonDeleteOU = New-Object System.Windows.Forms.Button
$script:buttonDeleteOU.Text = "删除OU组织"
$script:buttonDeleteOU.Width = 100
$script:buttonDeleteOU.BackColor = [System.Drawing.Color]::FromArgb(178, 34, 34)
$script:buttonDeleteOU.ForeColor = [System.Drawing.Color]::White
$script:buttonDeleteOU.FlatStyle = "Flat"
$script:buttonDeleteOU.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:buttonDeleteOU.Add_Click({ DeleteExistingOU })

$script:buttonRestrictLogin = New-Object System.Windows.Forms.Button
$script:buttonRestrictLogin.Text = "限制登录计算机"
$script:buttonRestrictLogin.Width = 120
$script:buttonRestrictLogin.BackColor = [System.Drawing.Color]::FromArgb(255, 140, 0)
$script:buttonRestrictLogin.ForeColor = [System.Drawing.Color]::White
$script:buttonRestrictLogin.FlatStyle = "Flat"
$script:buttonRestrictLogin.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:buttonRestrictLogin.Add_Click({ ShowRestrictLoginForm })

$script:buttonRestrictLogonTime = New-Object System.Windows.Forms.Button
$script:buttonRestrictLogonTime.Text = "限制登录时间"
$script:buttonRestrictLogonTime.Width = 110
$script:buttonRestrictLogonTime.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 150)
$script:buttonRestrictLogonTime.ForeColor = [System.Drawing.Color]::White
$script:buttonRestrictLogonTime.FlatStyle = "Flat"
$script:buttonRestrictLogonTime.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:buttonRestrictLogonTime.Add_Click({ ShowRestrictLogonTimeForm })

# 将所有按钮添加到同一个面板
$script:ouButtonPanel.Controls.Add($script:buttonSwitchOU)
$script:ouButtonPanel.Controls.Add($script:buttonCreateOU)
$script:ouButtonPanel.Controls.Add($script:buttonRenameOU)
$script:ouButtonPanel.Controls.Add($script:buttonDeleteOU)

$script:ouButtonPanel.Controls.Add($script:buttonRestrictLogin)
$script:ouButtonPanel.Controls.Add($script:buttonRestrictLogonTime)


# 只需要将面板添加到表格一次（选择合适的单元格，比如1,3）
$script:connectionTable.Controls.Add($script:ouButtonPanel, 1, 3)


# 将表格添加到连接面板
$script:connectionPanel.Controls.Add($script:connectionTable)

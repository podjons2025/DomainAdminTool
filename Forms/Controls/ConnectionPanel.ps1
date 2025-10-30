<# 
�������������� 
#>

$script:connectionPanel = New-Object System.Windows.Forms.GroupBox
$script:connectionPanel.Text = "�����������"
$script:connectionPanel.Dock = "Fill"
$script:connectionPanel.Padding = 5

# ��������񲼾�
$script:connectionTable = New-Object System.Windows.Forms.TableLayoutPanel
$script:connectionTable.Dock = "Fill"
$script:connectionTable.RowCount = 4 
$script:connectionTable.ColumnCount = 4
$script:connectionTable.Padding = 5
# �����б���
$script:connectionTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
$script:connectionTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
$script:connectionTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
$script:connectionTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
# �в���
$script:connectionTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100)))
$script:connectionTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$script:connectionTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100)))
$script:connectionTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

# 1. ��ص�ַ������
$script:labelDomain = New-Object System.Windows.Forms.Label
$script:labelDomain.Text = "��ص�ַ:"
$script:labelDomain.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:connectionTable.Controls.Add($script:labelDomain, 0, 0)

$script:comboDomain = New-Object System.Windows.Forms.ComboBox
$script:comboDomain.Dock = "Fill"
$script:comboDomain.DropDownStyle = "DropDownList"
$script:comboDomain.DisplayMember = "Name"
$script:comboDomain.ValueMember = "Server"
$script:comboDomain.Items.AddRange(@(	
    [PSCustomObject]@{Name = "��أ����ݣ�- serverAD.abc.com"; Server = "serverAD.abc.com"; SystemAccount= "abc\admin"; Password = "Abc123456"},
    [PSCustomObject]@{Name = "��أ��Ϻ���- abc03.abc01.com"; Server = "abc03.abc01.com"; SystemAccount= "abc01\administrator"; Password = "Password123"},
    [PSCustomObject]@{Name = "������أ�������- serverAD3.abc03.com"; Server = "serverAD3.abc03.com"; SystemAccount= "abc03\admin"; Password = ""}		
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

# 2. ����Ա�˺�
$script:labelAdmin = New-Object System.Windows.Forms.Label
$script:labelAdmin.Text = "����Ա�˺�:"
$script:labelAdmin.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:connectionTable.Controls.Add($script:labelAdmin, 2, 0)

$script:textAdmin = New-Object System.Windows.Forms.TextBox
$script:textAdmin.Text = $script:comboDomain.SelectedItem.SystemAccount 
$script:textAdmin.Dock = "Fill"
$script:connectionTable.Controls.Add($script:textAdmin, 3, 0)

# 3. ���������
$script:labelPassword = New-Object System.Windows.Forms.Label
$script:labelPassword.Text = "����:"
$script:labelPassword.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:connectionTable.Controls.Add($script:labelPassword, 0, 1)

$script:textPassword = New-Object System.Windows.Forms.TextBox
$script:textPassword.PasswordChar = '*'
$script:textPassword.Text = $script:comboDomain.SelectedItem.Password
$script:textPassword.Dock = "Fill"

# �����갴���¼� - ��ʾ����
$script:textPassword.Add_MouseDown({
    $this.PasswordChar = $null  # ȡ���������룬��ʾ����
})

# �������ͷ��¼� - �ָ�����
$script:textPassword.Add_MouseUp({
    $this.PasswordChar = '*'    # �ָ���������
})

# �������뿪�¼� - ȷ���뿪ʱ�ָ�����
$script:textPassword.Add_MouseLeave({
    $this.PasswordChar = '*'    # �ָ���������
})

$script:connectionTable.Controls.Add($script:textPassword, 1, 1)

# 4. OU��֯
$script:labelOU = New-Object System.Windows.Forms.Label
$script:labelOU.Text = "OU��֯:"
$script:labelOU.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:connectionTable.Controls.Add($script:labelOU, 0, 2)  # ��3�е�1��

$script:textOU = New-Object System.Windows.Forms.TextBox
$script:textOU.Dock = "Fill"
$script:textOU.ReadOnly = $true
$script:connectionTable.Controls.Add($script:textOU, 1, 2)  # ��3�е�2��

# 5. ����/�Ͽ���ť���
$script:buttonPanel = New-Object System.Windows.Forms.Panel
$script:buttonPanel.Dock = "Fill"
$script:buttonPanel.Padding = 5

$script:buttonConnect = New-Object System.Windows.Forms.Button
$script:buttonConnect.Text = "�������"
$script:buttonConnect.Location = New-Object System.Drawing.Point(5, 5)
$script:buttonConnect.Width = 80
$script:buttonConnect.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
$script:buttonConnect.ForeColor = [System.Drawing.Color]::White
$script:buttonConnect.FlatStyle = "Flat"
$script:buttonConnect.Add_Click({ ConnectToDomain })  # ����DomainOperations.ps1

$script:buttonDisconnect = New-Object System.Windows.Forms.Button
$script:buttonDisconnect.Text = "�˳�����"
$script:buttonDisconnect.Location = New-Object System.Drawing.Point(115, 5)
$script:buttonDisconnect.Width = 80
$script:buttonDisconnect.BackColor = [System.Drawing.Color]::FromArgb(169, 169, 169)
$script:buttonDisconnect.ForeColor = [System.Drawing.Color]::White
$script:buttonDisconnect.FlatStyle = "Flat"
$script:buttonDisconnect.Enabled = $false
$script:buttonDisconnect.Add_Click({ DisconnectFromDomain })  # ����DomainOperations.ps1

$script:buttonPanel.Controls.Add($script:buttonConnect)
$script:buttonPanel.Controls.Add($script:buttonDisconnect)
$script:connectionTable.Controls.Add($script:buttonPanel, 3, 1)  # �����������е��Ҳ�




# 6. OU������ť���
$script:ouButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$script:ouButtonPanel.Dock = "Fill"  # �������������
$script:ouButtonPanel.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)  # ����10px������5px�ڱ߾࣬�Ż��Ӿ�
$script:ouButtonPanel.FlowDirection = "LeftToRight"  # ��ť����������
$script:ouButtonPanel.WrapContents = $false  # ��ֹ��ť���У�����һ����ʾ
$script:ouButtonPanel.AutoScroll = $false  # ���������
$script:ouButtonPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)  # �������屳��ɫһ�£��Ӿ�ͳһ

$script:buttonSwitchOU = New-Object System.Windows.Forms.Button
$script:buttonSwitchOU.Text = "�л�OU��֯"
$script:buttonSwitchOU.Width = 100
$script:buttonSwitchOU.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
$script:buttonSwitchOU.ForeColor = [System.Drawing.Color]::White
$script:buttonSwitchOU.FlatStyle = "Flat"
$script:buttonSwitchOU.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)  # ��ť���
$script:buttonSwitchOU.Add_Click({ SwitchOU })

$script:buttonCreateOU = New-Object System.Windows.Forms.Button
$script:buttonCreateOU.Text = "�½�OU��֯"
$script:buttonCreateOU.Width = 100
$script:buttonCreateOU.BackColor = [System.Drawing.Color]::FromArgb(128, 0, 128)
$script:buttonCreateOU.ForeColor = [System.Drawing.Color]::White
$script:buttonCreateOU.FlatStyle = "Flat"
$script:buttonCreateOU.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:buttonCreateOU.Add_Click({ CreateNewOU })

$script:buttonRenameOU = New-Object System.Windows.Forms.Button
$script:buttonRenameOU.Text = "������OU��֯"
$script:buttonRenameOU.Width = 100
$script:buttonRenameOU.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)  # ����ɫ
$script:buttonRenameOU.ForeColor = [System.Drawing.Color]::White
$script:buttonRenameOU.FlatStyle = "Flat"
$script:buttonRenameOU.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:buttonRenameOU.Add_Click({ RenameExistingOU })

$script:buttonDeleteOU = New-Object System.Windows.Forms.Button
$script:buttonDeleteOU.Text = "ɾ��OU��֯"
$script:buttonDeleteOU.Width = 100
$script:buttonDeleteOU.BackColor = [System.Drawing.Color]::FromArgb(178, 34, 34)
$script:buttonDeleteOU.ForeColor = [System.Drawing.Color]::White
$script:buttonDeleteOU.FlatStyle = "Flat"
$script:buttonDeleteOU.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:buttonDeleteOU.Add_Click({ DeleteExistingOU })

$script:buttonRestrictLogin = New-Object System.Windows.Forms.Button
$script:buttonRestrictLogin.Text = "���Ƶ�¼�����"
$script:buttonRestrictLogin.Width = 120
$script:buttonRestrictLogin.BackColor = [System.Drawing.Color]::FromArgb(255, 140, 0)
$script:buttonRestrictLogin.ForeColor = [System.Drawing.Color]::White
$script:buttonRestrictLogin.FlatStyle = "Flat"
$script:buttonRestrictLogin.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:buttonRestrictLogin.Add_Click({ ShowRestrictLoginForm })

$script:buttonRestrictLogonTime = New-Object System.Windows.Forms.Button
$script:buttonRestrictLogonTime.Text = "���Ƶ�¼ʱ��"
$script:buttonRestrictLogonTime.Width = 110
$script:buttonRestrictLogonTime.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 150)
$script:buttonRestrictLogonTime.ForeColor = [System.Drawing.Color]::White
$script:buttonRestrictLogonTime.FlatStyle = "Flat"
$script:buttonRestrictLogonTime.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:buttonRestrictLogonTime.Add_Click({ ShowRestrictLogonTimeForm })

# �����а�ť��ӵ�ͬһ�����
$script:ouButtonPanel.Controls.Add($script:buttonSwitchOU)
$script:ouButtonPanel.Controls.Add($script:buttonCreateOU)
$script:ouButtonPanel.Controls.Add($script:buttonRenameOU)
$script:ouButtonPanel.Controls.Add($script:buttonDeleteOU)

$script:ouButtonPanel.Controls.Add($script:buttonRestrictLogin)
$script:ouButtonPanel.Controls.Add($script:buttonRestrictLogonTime)


# ֻ��Ҫ�������ӵ����һ�Σ�ѡ����ʵĵ�Ԫ�񣬱���1,3��
$script:connectionTable.Controls.Add($script:ouButtonPanel, 1, 3)


# �������ӵ��������
$script:connectionPanel.Controls.Add($script:connectionTable)

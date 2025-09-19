<# 
用户管理面板（左侧列表+右侧操作）
#>

$script:userManagementPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:userManagementPanel.Dock = "Fill"
$script:userManagementPanel.ColumnCount = 2
$script:userManagementPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$script:userManagementPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

# ---------------------- 左侧：用户列表 ----------------------
$script:userListPanel = New-Object System.Windows.Forms.GroupBox
$script:userListPanel.Text = "账号列表"
$script:userListPanel.Dock = "Fill"
$script:userListPanel.Padding = 10

$script:userListTable = New-Object System.Windows.Forms.TableLayoutPanel
$script:userListTable.Dock = "Fill"
$script:userListTable.RowCount = 2
$script:userListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$script:userListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

# 搜索框
$script:searchPanel = New-Object System.Windows.Forms.Panel
$script:searchPanel.Dock = "Fill"

$script:labelSearch = New-Object System.Windows.Forms.Label
$script:labelSearch.Text = "搜索账号:"
$script:labelSearch.Location = New-Object System.Drawing.Point(5, 10)
$script:labelSearch.AutoSize = $true

$script:textSearch = New-Object System.Windows.Forms.TextBox
$script:textSearch.Location = New-Object System.Drawing.Point(80, 7)
$script:textSearch.Width = 300
$script:textSearch.Add_TextChanged({ FilterUserList $script:textSearch.Text })  # 来自Helpers.ps1

$script:searchPanel.Controls.Add($script:labelSearch)
$script:searchPanel.Controls.Add($script:textSearch)
$script:userListTable.Controls.Add($script:searchPanel, 0, 0)

# 用户DataGridView
$script:userDataGridView = New-Object System.Windows.Forms.DataGridView
$script:userDataGridView.Dock = "Fill"
$script:userDataGridView.SelectionMode = "FullRowSelect"
$script:userDataGridView.MultiSelect = $true
$script:userDataGridView.ReadOnly = $false
$script:userDataGridView.AutoGenerateColumns = $false
$script:userDataGridView.AllowUserToAddRows = $false
$script:userDataGridView.RowHeadersVisible = $false
$script:userDataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 255)
$script:userDataGridView.ColumnHeadersHeightSizeMode = "AutoSize"
$script:userDataGridView.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter

# 列定义
$script:colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colDisplayName.HeaderText = "姓名"
$script:colDisplayName.DataPropertyName = "DisplayName"
$script:colDisplayName.Width = 100
$script:colDisplayName.ReadOnly = $true
$script:colDisplayName.Name = "DisplayName"

$script:colSamAccountName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colSamAccountName.HeaderText = "账号"
$script:colSamAccountName.DataPropertyName = "SamAccountName"
$script:colSamAccountName.Width = 110
$script:colSamAccountName.ReadOnly = $true
$script:colSamAccountName.Name = "SamAccountName" 

$script:colGroups = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colGroups.HeaderText = "所属组"
$script:colGroups.FillWeight = 150
$script:colGroups.ReadOnly = $true
$script:colGroups.DataPropertyName = "MemberOf"


$script:colEnabled = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$script:colEnabled.HeaderText = "账号状态"
$script:colEnabled.DataPropertyName = "Enabled"
$script:colEnabled.Width = 120
$script:colEnabled.ReadOnly = $false
$script:colEnabled.FalseValue = $false
$script:colEnabled.TrueValue = $true
$script:colEnabled.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
$script:colEnabled.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$script:colEnabled.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$script:colEnabled.Name = "Enabled"

$script:colLocked = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colLocked.HeaderText = "已锁定"
$script:colLocked.DataPropertyName = "AccountLockout"
$script:colLocked.Width = 80
$script:colLocked.ReadOnly = $true
$script:colLocked.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter

$script:colExpirationDate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colExpirationDate.HeaderText = "到期日期"
$script:colExpirationDate.DataPropertyName = "AccountExpirationDate"
$script:colExpirationDate.Width = 120
$script:colExpirationDate.ReadOnly = $true 
$script:colExpirationDate.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$script:colExpirationDate.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$script:colExpirationDate.Name = "AccountExpirationDate"

$script:colEmail = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colEmail.HeaderText = "邮箱"
$script:colEmail.DataPropertyName = "EmailAddress"
$script:colEmail.Width = 180
$script:colEmail.ReadOnly = $true

$script:colPhone = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colPhone.HeaderText = "电话"
$script:colPhone.DataPropertyName = "TelePhone"
$script:colPhone.Width = 100
$script:colPhone.ReadOnly = $true

$script:colDescription = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colDescription.HeaderText = "描述"
$script:colDescription.DataPropertyName = "Description"
$script:colDescription.Width = 180
$script:colDescription.ReadOnly = $true

$script:userDataGridView.Columns.AddRange($script:colDisplayName, $script:colSamAccountName, $script:colGroups, $script:colEnabled, $script:colLocked, $script:colExpirationDate, $script:colEmail, $script:colPhone, $script:colDescription)

# 单元格格式化事件
$script:userDataGridView.Add_CellFormatting({
    param($sender, $e)
    if ($sender.Columns[$e.ColumnIndex].DataPropertyName -eq "AccountLockout") {
        $rawValue = $e.Value
        if ($rawValue -eq $null) { $e.Value = "否"; $e.FormattingApplied = $true; return }
        $isLocked = switch ($rawValue.GetType().Name) {
            "String" { $rawValue -eq "True" }
            "Boolean" { [bool]$rawValue }
            default { $false }
        }
        $e.Value = if ($isLocked) { "是" } else { "否" }
        $e.FormattingApplied = $true
    }
	
    # 处理“时间期限”列（核心：读取到的时间减1天再对比）
    if ($sender.Columns[$e.ColumnIndex].DataPropertyName -eq "AccountExpirationDate") {
        $rawValue = $e.Value
        $currentDate = Get-Date -Date (Get-Date).Date  # 当前日期（仅年月日，时间00:00:00）

        # 1. 永不过期（AD未设置过期时间）
        if ($rawValue -eq $null -or $rawValue -is [DBNull]) {
            $e.Value = "永不过期"
            $e.CellStyle.ForeColor = [System.Drawing.Color]::Black
        }
        # 2. 有过期时间：减1天后再判断
        elseif ($rawValue -is [DateTime]) {
            # 关键：读取到的AD时间减1天，仅保留日期部分
            $adjustedExpiryDate = $rawValue.AddDays(-1).Date  # 减1天 + 清除时间
            
            # 已过期：调整后的日期 < 当前日期
            if ($adjustedExpiryDate -lt $currentDate) {
                $e.Value = "已过期"
                $e.CellStyle.ForeColor = [System.Drawing.Color]::Red
            }
            # 未过期：显示调整后的日期（与AD实际设置的日期一致）
            else {
                $e.Value = $adjustedExpiryDate.ToString("yyyy-MM-dd")
                $e.CellStyle.ForeColor = [System.Drawing.Color]::Black
            }
        }
        # 3. 异常值处理
        else {
            $e.Value = "未知"
            $e.CellStyle.ForeColor = [System.Drawing.Color]::Gray
        }

        $e.FormattingApplied = $true 
    }	
})




# 账号状态列绘制事件
$script:userDataGridView.Add_CellPainting({
    param($sender, $e)
    if ($e.ColumnIndex -ge 0 -and $sender.Columns[$e.ColumnIndex].Name -eq "Enabled" -and $e.RowIndex -ge 0) {
        $e.PaintBackground($e.CellBounds, $true)
        $e.PaintContent($e.CellBounds)
        $cellValue = $sender.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value
        $isEnabled = if ($cellValue -eq $null) { $false } else { [bool]$cellValue }
        $statusText = if ($isEnabled) { "启用" } else { "禁用" }
        $statusColor = if ($isEnabled) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red }
        
        $textFormat = [System.Drawing.StringFormat]::new()
        $textFormat.Alignment = [System.Drawing.StringAlignment]::Near
        $textFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
        $textRect = [System.Drawing.RectangleF]::new($e.CellBounds.X + 25, $e.CellBounds.Y, $e.CellBounds.Width - 30, $e.CellBounds.Height)
        $font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $e.Graphics.DrawString($statusText, $font, [System.Drawing.SolidBrush]$statusColor, $textRect, $textFormat)
        
        $textFormat.Dispose()
        $font.Dispose()
        $e.Handled = $true
    }
})

# 账号状态切换事件（启用/禁用）
$script:userDataGridView.Add_CellContentClick({
    if ($_.ColumnIndex -eq $script:colEnabled.Index -and $_.RowIndex -ge 0) {
        ToggleUserEnabled $_.RowIndex  # 来自UserOperations.ps1
    }
})

# 用户选择变化事件
$script:userDataGridView.Add_SelectionChanged({
    if ($script:userDataGridView.SelectedRows.Count -gt 0) {
        $user = $script:userDataGridView.SelectedRows[0].DataBoundItem
        $script:textCnName.Text = $user.DisplayName 
        $script:textPinyin.Text = $user.SamAccountName
        $script:textEmail.Text = $user.EmailAddress
		$script:textPhone.Text = $user.TelePhone
        $script:textDescription.Text = $user.Description
        $script:textNewPassword.Text = ""
        $script:textConfirmPassword.Text = ""
    }
})

$script:userListTable.Controls.Add($script:userDataGridView, 0, 1)
$script:userListPanel.Controls.Add($script:userListTable)
$script:userManagementPanel.Controls.Add($script:userListPanel, 0, 0)

# ---------------------- 右侧：用户操作面板 ----------------------

$script:userOperationPanel = New-Object System.Windows.Forms.GroupBox
$script:userOperationPanel.Text = "账号操作"
$script:userOperationPanel.Dock = "Fill"
$script:userOperationPanel.Padding = New-Object System.Windows.Forms.Padding(20, 10, 10, 10)  # 左侧加宽内边距，避免标签拥挤

$script:operationTable = New-Object System.Windows.Forms.TableLayoutPanel
$script:operationTable.Dock = "Fill"
$script:operationTable.RowCount = 7  # 7行结构（姓名、登录+前缀、邮箱、描述、到期、新密码、确认密码）
$script:operationTable.ColumnCount = 2  # 2列：标签列 + 控件列
$script:operationTable.Padding = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:operationTable.CellBorderStyle = "None"

# 统一行高为35px，保证垂直间距均匀
for ($i=0; $i -lt $script:operationTable.RowCount; $i++) {
    $script:operationTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
}
$script:operationTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))  # 标签列固定宽度
$script:operationTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # 控件列占满剩余宽度


# ---------- 1. 姓名 ----------
$script:labelCnName = New-Object System.Windows.Forms.Label
$script:labelCnName.Text = "姓名:"
$script:labelCnName.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # 标签文字右居中
$script:labelCnName.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)  # 右侧增加间距，与输入框更协调
$script:operationTable.Controls.Add($script:labelCnName, 0, 0)

$script:textCnName = New-Object System.Windows.Forms.TextBox
$script:textCnName.Dock = "Fill"
$script:textCnName.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)
$script:textCnName.Add_TextChanged({ ConvertToPinyin })  # 来自PinyinConverter.ps1
$script:operationTable.Controls.Add($script:textCnName, 1, 0)


# ---------- 2. 登录账号 + 账号前缀 ----------
$script:labelPinyin = New-Object System.Windows.Forms.Label
$script:labelPinyin.Text = "登录账号:"
$script:labelPinyin.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # 与“姓名”标签对齐
$script:labelPinyin.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelPinyin, 0, 1)

$script:accountSubPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:accountSubPanel.Dock = "Fill"
$script:accountSubPanel.ColumnCount = 3  # 3列：登录输入框、前缀标签、前缀输入框
$script:accountSubPanel.RowCount = 1
$script:accountSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
$script:accountSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 80)))
$script:accountSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$script:accountSubPanel.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)

$script:textPinyin = New-Object System.Windows.Forms.TextBox
$script:textPinyin.Dock = "Fill"
$script:textPinyin.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)  # 与前缀标签留间距
$script:accountSubPanel.Controls.Add($script:textPinyin, 0, 0)

$script:labelPrefix = New-Object System.Windows.Forms.Label
$script:labelPrefix.Text = "账号前缀:"
$script:labelPrefix.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:labelPrefix.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)
$script:accountSubPanel.Controls.Add($script:labelPrefix, 1, 0)

$script:textPrefix = New-Object System.Windows.Forms.TextBox
$script:textPrefix.Text = "abc_"
$script:textPrefix.Dock = "Fill"
$script:textPrefix.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$script:textPrefix.Add_TextChanged({ ConvertToPinyin })  # 来自PinyinConverter.ps1
$script:accountSubPanel.Controls.Add($script:textPrefix, 2, 0)

$script:operationTable.Controls.Add($script:accountSubPanel, 1, 1)


# ---------- 3. 邮箱 + 电话 ----------
$script:labelEmail = New-Object System.Windows.Forms.Label
$script:labelEmail.Text = "邮箱:"
$script:labelEmail.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight 
$script:labelEmail.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelEmail, 0, 2)

# 右侧创建“邮箱输入框+电话标签+电话输入框”的子面板
$script:contactSubPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:contactSubPanel.Dock = "Fill"
$script:contactSubPanel.ColumnCount = 4
$script:contactSubPanel.RowCount = 1
$script:contactSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
$script:contactSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 80)))
$script:contactSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$script:contactSubPanel.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)

# 1. 邮箱输入框
$script:textEmail = New-Object System.Windows.Forms.TextBox
$script:textEmail.Dock = "Fill"
$script:textEmail.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0) 
$script:contactSubPanel.Controls.Add($script:textEmail, 0, 0)

# 2. 电话标签
$script:labelPhone = New-Object System.Windows.Forms.Label
$script:labelPhone.Text = "联系电话:"
$script:labelPhone.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight 
$script:labelPhone.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0) 
$script:contactSubPanel.Controls.Add($script:labelPhone, 1, 0)

# 3. 电话输入框
$script:textPhone = New-Object System.Windows.Forms.TextBox
$script:textPhone.Dock = "Fill"
$script:textPhone.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$script:contactSubPanel.Controls.Add($script:textPhone, 2, 0)

# 将子面板添加到主表格
$script:operationTable.Controls.Add($script:contactSubPanel, 1, 2)


# ---------- 4. 描述 ----------
$script:labelDescription = New-Object System.Windows.Forms.Label
$script:labelDescription.Text = "描述:"
$script:labelDescription.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:labelDescription.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelDescription, 0, 3)

$script:textDescription = New-Object System.Windows.Forms.TextBox
$script:textDescription.Dock = "Fill"
$script:textDescription.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)
$script:operationTable.Controls.Add($script:textDescription, 1, 3)


# ---------- 5. 到期日期 ----------
$script:labelExpiry = New-Object System.Windows.Forms.Label
$script:labelExpiry.Text = "到期日期:"
$script:labelExpiry.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:labelExpiry.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelExpiry, 0, 4)

$script:expiryPanel = New-Object System.Windows.Forms.Panel
$script:expiryPanel.Dock = "Fill"
$script:expiryPanel.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)

# 日期选择器
$script:dateExpiry = New-Object System.Windows.Forms.DateTimePicker
$script:dateExpiry.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$script:dateExpiry.Location = New-Object System.Drawing.Point(0, 5)
$script:dateExpiry.Width = 120
$script:dateExpiry.Height = 22
$script:dateExpiry.Enabled = $false

# 永不过期复选框
$script:chkNeverExpire = New-Object System.Windows.Forms.CheckBox
$script:chkNeverExpire.Text = "永不过期"
$script:chkNeverExpire.Location = New-Object System.Drawing.Point(135, 5)
$script:chkNeverExpire.Width = 80
$script:chkNeverExpire.Checked = $true
$script:chkNeverExpire.Add_CheckedChanged({ $script:dateExpiry.Enabled = -not $this.Checked })

# 初始密码标签（位于“录入密码”按钮左侧）
$script:labelPassword = New-Object System.Windows.Forms.Label
$script:labelPassword.Text = "(密码：Password@001)"
$script:labelPassword.Height = 25
$script:labelPassword.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$script:labelPassword.Location = New-Object System.Drawing.Point(213, 5)  # 调整位置到按钮左侧
$script:labelPassword.AutoSize = $true  # 自动适应文本宽度


# 初始密码按钮（左移，与初始密码标签对齐）
$script:buttonCreatePassword = New-Object System.Windows.Forms.Button
$script:buttonCreatePassword.Text = "初始密码"
$script:buttonCreatePassword.Location = New-Object System.Drawing.Point(340, 5)  # 位于初始密码标签右侧
$script:buttonCreatePassword.Width = 80
$script:buttonCreatePassword.Height = 25
$script:buttonCreatePassword.BackColor = [System.Drawing.Color]::FromArgb(100, 150, 250)
$script:buttonCreatePassword.ForeColor = [System.Drawing.Color]::White
$script:buttonCreatePassword.FlatStyle = "Flat"
# 点击事件：填入初始密码
$script:buttonCreatePassword.Add_Click({ 
    $script:textNewPassword.Text = "Password@001"
    $script:textConfirmPassword.Text = "Password@001"
})

# 将所有控件添加到面板
$script:expiryPanel.Controls.Add($script:dateExpiry)
$script:expiryPanel.Controls.Add($script:chkNeverExpire)
$script:expiryPanel.Controls.Add($script:labelPassword)  # 新增初始密码标签
$script:expiryPanel.Controls.Add($script:buttonCreatePassword)

$script:operationTable.Controls.Add($script:expiryPanel, 1, 4)


# 新密码（自动填入初始密码）
$script:labelNewPassword = New-Object System.Windows.Forms.Label
$script:labelNewPassword.Text = "新密码:"
$script:labelNewPassword.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:labelNewPassword.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelNewPassword, 0, 5)

$script:textNewPassword = New-Object System.Windows.Forms.TextBox
$script:textNewPassword.PasswordChar = '*'
$script:textNewPassword.Dock = "Fill"
$script:textNewPassword.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)
$script:textNewPassword.Text = ""
$script:operationTable.Controls.Add($script:textNewPassword, 1, 5)


#确认密码（自动填入初始密码，增加底部留空）
$script:labelConfirmPassword = New-Object System.Windows.Forms.Label
$script:labelConfirmPassword.Text = "确认密码:"
$script:labelConfirmPassword.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:labelConfirmPassword.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelConfirmPassword, 0, 6)

$script:textConfirmPassword = New-Object System.Windows.Forms.TextBox
$script:textConfirmPassword.PasswordChar = '*'
$script:textConfirmPassword.Dock = "Fill"
$script:textConfirmPassword.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 10)
$script:textConfirmPassword.Text = ""
$script:operationTable.Controls.Add($script:textConfirmPassword, 1, 6)


# 统一设置文本框最小高度
foreach ($tb in $script:operationTable.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] }) {
    $tb.MinimumSize = New-Object System.Drawing.Size(0, 22)
}

$script:userOperationPanel.Controls.Add($script:operationTable)
$script:userManagementPanel.Controls.Add($script:userOperationPanel, 1, 0)



# 中间操作按钮（复用给主面板）
$script:middleButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$script:middleButtonPanel.Dock = "Fill"
$script:middleButtonPanel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
$script:middleButtonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
$script:middleButtonPanel.WrapContents = $false
$script:middleButtonPanel.Height = 60

$script:buttonCreate = New-Object System.Windows.Forms.Button
$script:buttonCreate.Text = "新建账号"
$script:buttonCreate.Width = 80
$script:buttonCreate.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonCreate.BackColor = [System.Drawing.Color]::ForestGreen
$script:buttonCreate.ForeColor = [System.Drawing.Color]::White
$script:buttonCreate.FlatStyle = "Flat"
$script:buttonCreate.Add_Click({ CreateNewUser })  # 来自UserOperations.ps1

$script:buttonChangePassword = New-Object System.Windows.Forms.Button
$script:buttonChangePassword.Text = "修改密码"
$script:buttonChangePassword.Width = 80
$script:buttonChangePassword.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonChangePassword.BackColor = [System.Drawing.Color]::Orange
$script:buttonChangePassword.ForeColor = [System.Drawing.Color]::White
$script:buttonChangePassword.FlatStyle = "Flat"
$script:buttonChangePassword.Add_Click({ ChangeUserPassword })  # 来自UserOperations.ps1

$script:buttonModifyUser = New-Object System.Windows.Forms.Button
$script:buttonModifyUser.Text = "修改信息"
$script:buttonModifyUser.Width = 80
$script:buttonModifyUser.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonModifyUser.BackColor = [System.Drawing.Color]::DarkCyan
$script:buttonModifyUser.ForeColor = [System.Drawing.Color]::White
$script:buttonModifyUser.FlatStyle = "Flat"
$script:buttonModifyUser.Add_Click({ ModifyUserAccount })  # 来自UserOperations.ps1


$script:buttonUnlock = New-Object System.Windows.Forms.Button
$script:buttonUnlock.Text = "解锁账号"
$script:buttonUnlock.Width = 80
$script:buttonUnlock.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonUnlock.BackColor = [System.Drawing.Color]::DarkOrchid
$script:buttonUnlock.ForeColor = [System.Drawing.Color]::White
$script:buttonUnlock.FlatStyle = "Flat"
$script:buttonUnlock.Add_Click({ UnlockUserAccount })  # 来自UserOperations.ps1

$script:buttonRefresh = New-Object System.Windows.Forms.Button
$script:buttonRefresh.Text = "刷新列表"
$script:buttonRefresh.Width = 80
$script:buttonRefresh.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonRefresh.BackColor = [System.Drawing.Color]::SteelBlue
$script:buttonRefresh.ForeColor = [System.Drawing.Color]::White
$script:buttonRefresh.FlatStyle = "Flat"
$script:buttonRefresh.Add_Click({ LoadUserList; LoadGroupList })  # 来自User/GroupOperations.ps1

$script:buttonRename = New-Object System.Windows.Forms.Button
$script:buttonRename.Text = "重命名账号"
$script:buttonRename.Width = 80
$script:buttonRename.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonRename.BackColor = [System.Drawing.Color]::MediumPurple
$script:buttonRename.ForeColor = [System.Drawing.Color]::White
$script:buttonRename.FlatStyle = "Flat"
$script:buttonRename.Add_Click({ RenameUserAccount })  # 来自UserOperations.ps1

$script:buttonDelete = New-Object System.Windows.Forms.Button
$script:buttonDelete.Text = "删除账号"
$script:buttonDelete.Width = 80
$script:buttonDelete.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonDelete.BackColor = [System.Drawing.Color]::Crimson
$script:buttonDelete.ForeColor = [System.Drawing.Color]::White
$script:buttonDelete.FlatStyle = "Flat"
$script:buttonDelete.Add_Click({ DeleteUserAccount })  # 来自UserOperations.ps1

# 批量导入CSV按钮
$script:buttonImportCSV = New-Object System.Windows.Forms.Button
$script:buttonImportCSV.Text = "导入CSV批量创建"
$script:buttonImportCSV.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonImportCSV.Width = 110
$script:buttonImportCSV.BackColor = [System.Drawing.Color]::FromArgb(30, 100, 120)
$script:buttonImportCSV.ForeColor = [System.Drawing.Color]::White
$script:buttonImportCSV.FlatStyle = "Flat"
$script:buttonImportCSV.Add_Click({ImportCSVAndCreateUsers})  # 来自importExportUsers.ps1

# 导出CSV按钮
$script:buttonExportCSV = New-Object System.Windows.Forms.Button
$script:buttonExportCSV.Text = "导出CSV"
$script:buttonExportCSV.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonExportCSV.Width = 90
$script:buttonExportCSV.BackColor = [System.Drawing.Color]::FromArgb(150, 120, 80)
$script:buttonExportCSV.ForeColor = [System.Drawing.Color]::White
$script:buttonExportCSV.FlatStyle = "Flat"
$script:buttonExportCSV.Add_Click({ExportCSVUsers})   # 来自importExportUsers.ps1

$script:middleButtonPanel.Controls.AddRange(@($script:buttonDelete, `
$script:buttonChangePassword, $script:buttonModifyUser, $script:buttonUnlock,`
$script:buttonRefresh, $script:buttonRename, $script:buttonCreate, `
$script:buttonExportCSV, $script:buttonImportCSV
))




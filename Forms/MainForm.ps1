<# 
主窗体定义 
#>

$script:mainForm = New-Object System.Windows.Forms.Form
$script:mainForm.Text = "域控账号管理工具"
$script:mainForm.Size = New-Object System.Drawing.Size(1200, 900)
$script:mainForm.StartPosition = "CenterScreen"
#$script:mainForm.FormBorderStyle = "Fixed3D"
#$script:mainForm.MaximizeBox = $false
$script:mainForm.FormBorderStyle = "Sizable"
$script:mainForm.MaximizeBox = $true
$script:mainForm.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)

# 主布局面板（5行结构）
$script:mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:mainPanel.Dock = "Fill"
$script:mainPanel.RowCount = 5
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 180)))  # 增加域控连接面板高度
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 45)))   # 用户管理
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45)))  # 中间按钮
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 45)))   # 组管理
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))  # 状态显示

# 1. 添加域控连接面板
$script:mainPanel.Controls.Add($script:connectionPanel, 0, 0)

# 2. 添加用户管理面板
$script:mainPanel.Controls.Add($script:userManagementPanel, 0, 1)

# 3. 添加中间操作按钮
$script:mainPanel.Controls.Add($script:middleButtonPanel, 0, 2)

# 4. 添加组管理面板
$script:mainPanel.Controls.Add($script:groupManagementPanel, 0, 3)

# 5. 添加状态栏
$script:mainPanel.Controls.Add($script:statusOutputLabel, 0, 4)

# 将主面板添加到主窗体
$script:mainForm.Controls.Add($script:mainPanel)

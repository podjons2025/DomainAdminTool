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

# 主窗口大小变化事件（强制刷新所有父容器+延迟适配）
$script:mainForm.Add_SizeChanged({
    # 1. 强制刷新所有嵌套父面板（确保子控件尺寸同步更新）
    $script:mainPanel.PerformLayout()          # 最外层主面板
    $script:userManagementPanel.PerformLayout()# 用户管理面板
    $script:groupManagementPanel.PerformLayout()# 组管理面板
    $script:userListPanel.PerformLayout()      # 用户列表子面板
    $script:groupListPanel.PerformLayout()     # 组列表子面板
	$script:ouButtonPanel.PerformLayout()      # 刷新OU按钮面板布局

    # 2. 最大化/还原：延迟50ms（等待系统完成布局），最小化不处理
    if ($script:mainForm.WindowState -in [System.Windows.Forms.FormWindowState]::Maximized, [System.Windows.Forms.FormWindowState]::Normal) {
        Start-Sleep -Milliseconds 50  # 给系统足够时间更新控件尺寸
        Update-DynamicUserPageSize
        Update-DynamicGroupPageSize
    }
})

# DataGridView创建后立即计算首次动态行数（避免初始值为0）
$script:userDataGridView.PerformLayout()  # 确保DGV布局已初始化
Update-DynamicUserPageSize  # 提前计算动态行数

$script:userDataGridView.Add_SizeChanged({
    Start-Sleep -Milliseconds 50
    Update-DynamicUserPageSize
})

# DataGridView创建后立即计算首次动态行数（避免初始值为0）
$script:groupDataGridView.PerformLayout()  # 确保DGV布局已初始化
Update-DynamicGroupPageSize  # 提前计算动态行数

$script:groupDataGridView.Add_SizeChanged({
    Start-Sleep -Milliseconds 50
    Update-DynamicGroupPageSize
})

# 主布局面板（5行结构）
$script:mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:mainPanel.Dock = "Fill"
$script:mainPanel.RowCount = 6 
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 160)))  # 增加域控连接面板高度
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45)))  # 上层按钮
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 45)))   # 用户管理
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45)))  # 中间按钮
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 45)))   # 组管理
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))  # 状态显示

# 1. 域控连接面板（第0行）
$script:mainPanel.Controls.Add($script:connectionPanel, 0, 0)

# 2. OU操作按钮面板（第1行，独立一行，移出域控连接GroupBox）
$script:mainPanel.Controls.Add($script:ouButtonPanel, 0, 1)

# 3. 用户管理面板（第2行，原行索引后移1位）
$script:mainPanel.Controls.Add($script:userManagementPanel, 0, 2)

# 4. 中间操作按钮（第3行，原行索引后移1位）
$script:mainPanel.Controls.Add($script:middleButtonPanel, 0, 3)

# 5. 组管理面板（第4行，原行索引后移1位）
$script:mainPanel.Controls.Add($script:groupManagementPanel, 0, 4)

# 6. 状态栏（第5行，原行索引后移1位）
$script:mainPanel.Controls.Add($script:statusOutputLabel, 0, 5)

# 将主面板添加到主窗体
$script:mainForm.Controls.Add($script:mainPanel)

# 主窗口加载完成事件（新增：窗口显示后再计算初始行数）
$script:mainForm.Add_Load({
    Start-Sleep -Milliseconds 100  # 等待窗口完全渲染
    Update-DynamicUserPageSize
    Update-DynamicGroupPageSize
    $script:ouButtonPanel.PerformLayout()  # 刷新OU按钮面板
    # 强制刷新DataGridView，确保滚动条状态更新
    $script:userDataGridView.Refresh()
    $script:groupDataGridView.Refresh()
})

<# 
组管理面板（左侧列表+右侧操作）- 分页控件移至右下角版本
#>

$script:groupManagementPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:groupManagementPanel.Dock = "Fill"
$script:groupManagementPanel.ColumnCount = 2
$script:groupManagementPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$script:groupManagementPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

# 左侧：组列表 
$script:groupListPanel = New-Object System.Windows.Forms.GroupBox
$script:groupListPanel.Text = "组列表"
$script:groupListPanel.Dock = "Fill"
$script:groupListPanel.Padding = 10

$script:groupListTable = New-Object System.Windows.Forms.TableLayoutPanel
$script:groupListTable.Dock = "Fill"
$script:groupListTable.RowCount = 3  # 调整行顺序：搜索框→组DataGridView→分页面板
$script:groupListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))  # 搜索框（固定高度）
$script:groupListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # 组DataGridView（占满中间空间）
$script:groupListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))  # 分页面板（固定高度，位于底部）

# 组搜索框（保持原有逻辑，位置不变）
$script:groupSearchPanel = New-Object System.Windows.Forms.Panel
$script:groupSearchPanel.Dock = "Fill"

$script:labelGroupSearch = New-Object System.Windows.Forms.Label
$script:labelGroupSearch.Text = "搜索组:"
$script:labelGroupSearch.Location = New-Object System.Drawing.Point(5, 10)
$script:labelGroupSearch.AutoSize = $true

$script:textGroupSearch = New-Object System.Windows.Forms.TextBox
$script:textGroupSearch.Location = New-Object System.Drawing.Point(80, 7)
$script:textGroupSearch.Width = 300
# 修改搜索事件：过滤后触发分页（逻辑不变）
$script:textGroupSearch.Add_TextChanged({
    $filterText = $script:textGroupSearch.Text.ToLower()
    $script:filteredGroups.Clear()

    # 过滤逻辑（与原FilterGroupList一致）
    if ([string]::IsNullOrEmpty($filterText)) {
        $script:allGroups | ForEach-Object { $script:filteredGroups.Add($_) | Out-Null }
    } else {
        $script:allGroups | Where-Object {
            $_.Name.ToLower() -like "*$filterText*" -or
            $_.SamAccountName.ToLower() -like "*$filterText*" -or
            ( (-not [string]::IsNullOrEmpty($_.Description)) -and $_.Description.ToLower() -like "*$filterText*" )
        } | ForEach-Object { $script:filteredGroups.Add($_) | Out-Null }
    }

    # 重置分页状态并显示第一页
    $script:currentGroupPage = 1
    $script:totalGroupPages = Get-TotalPages -totalCount $script:filteredGroups.Count -pageSize $script:pageSize
    Show-GroupPage
})

$script:groupSearchPanel.Controls.Add($script:labelGroupSearch)
$script:groupSearchPanel.Controls.Add($script:textGroupSearch)
$script:groupListTable.Controls.Add($script:groupSearchPanel, 0, 0)  # 搜索框保留在第1行（索引0）

# ---------------------- 2. 组DataGridView（调整位置至中间，原分页面板位置） ----------------------
$script:groupDataGridView = New-Object System.Windows.Forms.DataGridView
$script:groupDataGridView.Dock = "Fill"
$script:groupDataGridView.SelectionMode = "FullRowSelect"
$script:groupDataGridView.MultiSelect = $true
$script:groupDataGridView.ReadOnly = $true
$script:groupDataGridView.AutoGenerateColumns = $false
$script:groupDataGridView.AllowUserToAddRows = $false
$script:groupDataGridView.RowHeadersVisible = $false
$script:groupDataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 255)
$script:groupDataGridView.ColumnHeadersHeightSizeMode = "AutoSize"

# 组列定义（保持不变）
$script:colGroupName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colGroupName.HeaderText = "组名称"
$script:colGroupName.DataPropertyName = "Name"
$script:colGroupName.Width = 150
$script:colGroupName.ReadOnly = $true

$script:colGroupSamAccountName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colGroupSamAccountName.HeaderText = "组账号"
$script:colGroupSamAccountName.DataPropertyName = "SamAccountName"
$script:colGroupSamAccountName.Width = 150
$script:colGroupSamAccountName.ReadOnly = $true

$script:colGroupDescription = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colGroupDescription.HeaderText = "描述"
$script:colGroupDescription.DataPropertyName = "Description"
$script:colGroupDescription.Width = 300
$script:colGroupDescription.ReadOnly = $true

$script:groupDataGridView.Columns.AddRange($script:colGroupName, $script:colGroupSamAccountName, $script:colGroupDescription)
$script:groupListTable.Controls.Add($script:groupDataGridView, 0, 1)  # 移至第2行（索引1），位于搜索框下方、分页上方

# ---------------------- 3. 组分页面板 ----------------------
$script:groupPaginationPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:groupPaginationPanel.Dock = "Fill"
$script:groupPaginationPanel.Visible = $false  # 默认隐藏（数据>10条时显示）
#$script:groupPaginationPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom  # 锚定右下角，窗口缩放时保持位置
$script:groupPaginationPanel.ColumnCount = 4  # 4列：上一页按钮、分页信息、下一页按钮
$script:groupPaginationPanel.RowCount = 1     # 1行
$script:groupPaginationPanel.Padding = New-Object System.Windows.Forms.Padding(0, 5, 10, 0)  # 右侧留空10px，避免贴边；顶部留空5px，优化垂直间距
$script:groupPaginationPanel.CellBorderStyle = "None"

# 列样式配置：按钮固定宽度，分页信息自适应文本宽度
$script:groupPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 70)))  # 1.上一页
$script:groupPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))    # 2.分页信息
$script:groupPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))    # 3.跳转控件（新增）
$script:groupPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 70)))  # 4.下一页

# 行样式：固定高度25px，匹配按钮高度
$script:groupPaginationPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 25)))

# 上一页按钮
$script:btnGroupPrev = New-Object System.Windows.Forms.Button
$script:btnGroupPrev.Text = "上一页"
$script:btnGroupPrev.Width = 65
$script:btnGroupPrev.Enabled = $false
$script:btnGroupPrev.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)  # 右侧留5px间距，与分页信息分隔
$script:btnGroupPrev.Add_Click({
    if ($script:currentGroupPage -gt 1) {
		$script:groupDefaultShowAll = $false  # 关闭默认全显
        $script:currentGroupPage--
        Show-GroupPage
    }
})
$script:groupPaginationPanel.Controls.Add($script:btnGroupPrev, 0, 0)  # 添加到第1列（索引0）

# 分页信息标签
$script:lblGroupPageInfo = New-Object System.Windows.Forms.Label
$script:lblGroupPageInfo.Text = "第 1 页 / 共 1 页（总计 0 条）"
$script:lblGroupPageInfo.AutoSize = $true
$script:lblGroupPageInfo.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 0)  # 顶部留5px，实现垂直居中；右侧留5px，与下一页按钮分隔
$script:groupPaginationPanel.Controls.Add($script:lblGroupPageInfo, 1, 0)  # 添加到第2列（索引1）

# 下一页按钮
$script:btnGroupNext = New-Object System.Windows.Forms.Button
$script:btnGroupNext.Text = "下一页"
$script:btnGroupNext.Width = 65
$script:btnGroupNext.Enabled = $false
$script:btnGroupNext.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$script:btnGroupNext.Add_Click({
    if ($script:currentGroupPage -lt $script:totalGroupPages) {
		$script:groupDefaultShowAll = $false  # 关闭默认全显
        $script:currentGroupPage++
        Show-GroupPage
    }
})
$script:groupPaginationPanel.Controls.Add($script:btnGroupNext, 2, 0)  # 添加到第3列（索引2）

# 将分页面板添加到groupListTable的第3行（索引2），位于DataGridView下方
$script:groupListTable.Controls.Add($script:groupPaginationPanel, 0, 2)

$script:groupListPanel.Controls.Add($script:groupListTable)
$script:groupManagementPanel.Controls.Add($script:groupListPanel, 0, 0)

# ---------------------- 新增：组分页跳转控件 ----------------------
$script:groupJumpPanel = New-Object System.Windows.Forms.Panel
$script:groupJumpPanel.AutoSize = $true  # 自适应内容宽度
$script:groupJumpPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)

# 跳转标签
$script:lblGroupJump = New-Object System.Windows.Forms.Label
$script:lblGroupJump.Text = "跳至："
$script:lblGroupJump.AutoSize = $true
$script:lblGroupJump.Location = New-Object System.Drawing.Point(10, 5)  # 垂直居中

# 页码输入框（限制数字输入）
$script:txtGroupJumpPage = New-Object System.Windows.Forms.TextBox
$script:txtGroupJumpPage.Width = 40  # 固定宽度
$script:txtGroupJumpPage.Location = New-Object System.Drawing.Point(52, 0)
$script:txtGroupJumpPage.MaxLength = 3  # 限制最大3位
# 只允许输入数字和退格键
$script:txtGroupJumpPage.Add_KeyPress({
    if (-not ([char]::IsDigit($_.KeyChar) -or $_.KeyChar -eq [char]8)) {
        $_.Handled = $true  # 阻止非数字输入
    }
})

# 跳转按钮
$script:btnGroupJump = New-Object System.Windows.Forms.Button
$script:btnGroupJump.Text = "跳转"
$script:btnGroupJump.Width = 50
$script:btnGroupJump.Location = New-Object System.Drawing.Point(100, 0)
$script:btnGroupJump.Add_Click({
    # 1. 输入验证
    $jumpPage = $script:txtGroupJumpPage.Text.Trim()
    if ([string]::IsNullOrEmpty($jumpPage)) {
        [System.Windows.Forms.MessageBox]::Show("请输入要跳转的页码！", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    # 关键修复：提前声明 $jumpPageInt 变量
    $jumpPageInt = 0
    if (-not [int]::TryParse($jumpPage, [ref]$jumpPageInt)) {
        [System.Windows.Forms.MessageBox]::Show("请输入有效的数字页码！", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    # 2. 页码范围验证
    if ($jumpPageInt -lt 1 -or $jumpPageInt -gt $script:totalGroupPages) {
        [System.Windows.Forms.MessageBox]::Show("页码超出范围！请输入 1 ~ $($script:totalGroupPages) 之间的页码", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    # 3. 执行跳转
	$script:groupDefaultShowAll = $false
    $script:currentGroupPage = $jumpPageInt
    Show-GroupPage
})

$script:lblGroupText = New-Object System.Windows.Forms.Label
$script:lblGroupText.Text = "(分页)"
$script:lblGroupText.AutoSize = $true
$script:lblGroupText.Location = New-Object System.Drawing.Point(150, 5)  # 垂直居中

# 将跳转控件添加到子面板
$script:groupJumpPanel.Controls.Add($script:lblGroupJump)
$script:groupJumpPanel.Controls.Add($script:txtGroupJumpPage)
$script:groupJumpPanel.Controls.Add($script:btnGroupJump)
$script:groupJumpPanel.Controls.Add($script:lblGroupText)
# 将子面板添加到分页面板第3列（索引2）
$script:groupPaginationPanel.Controls.Add($script:groupJumpPanel, 3, 0)

# ---------------------- 右侧：组操作面板 ----------------------
$script:groupOperationPanel = New-Object System.Windows.Forms.GroupBox
$script:groupOperationPanel.Text = "组操作"
$script:groupOperationPanel.Dock = "Fill"
$script:groupOperationPanel.Padding = 15

$script:groupOpTable = New-Object System.Windows.Forms.TableLayoutPanel
$script:groupOpTable.Dock = "Fill"
$script:groupOpTable.RowCount = 5
$script:groupOpTable.ColumnCount = 2
$script:groupOpTable.Padding = 10
$script:groupOpTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$script:groupOpTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$script:groupOpTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$script:groupOpTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 80)))
$script:groupOpTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$script:groupOpTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
$script:groupOpTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

# 1. 组名称
$script:labelGroupName = New-Object System.Windows.Forms.Label
$script:labelGroupName.Text = "组名称:"
$script:labelGroupName.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:groupOpTable.Controls.Add($script:labelGroupName, 0, 0)

$script:textGroupName = New-Object System.Windows.Forms.TextBox
$script:textGroupName.Dock = "Fill"
$script:groupOpTable.Controls.Add($script:textGroupName, 1, 0)

# 2. 组账号
$script:labelGroupSamAccount = New-Object System.Windows.Forms.Label
$script:labelGroupSamAccount.Text = "组账号:"
$script:labelGroupSamAccount.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:groupOpTable.Controls.Add($script:labelGroupSamAccount, 0, 1)

$script:textGroupSamAccount = New-Object System.Windows.Forms.TextBox
$script:textGroupSamAccount.Dock = "Fill"
$script:groupOpTable.Controls.Add($script:textGroupSamAccount, 1, 1)

# 3. 组描述
$script:labelGroupDescription = New-Object System.Windows.Forms.Label
$script:labelGroupDescription.Text = "组描述:"
$script:labelGroupDescription.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:groupOpTable.Controls.Add($script:labelGroupDescription, 0, 2)

$script:textGroupDescription = New-Object System.Windows.Forms.TextBox
$script:textGroupDescription.Dock = "Fill"
$script:groupOpTable.Controls.Add($script:textGroupDescription, 1, 2)

# 4. 组操作按钮
$script:groupButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$script:groupButtonsPanel.Dock = "Fill"
$script:groupButtonsPanel.Padding = New-Object System.Windows.Forms.Padding(-10, 10, 10, 10)
$script:groupButtonsPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$script:groupButtonsPanel.WrapContents = $false
$script:groupButtonsPanel.Height = 60

$script:buttonAddGroup = New-Object System.Windows.Forms.Button
$script:buttonAddGroup.Text = "新建组"
$script:buttonAddGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonAddGroup.Width = 70
$script:buttonAddGroup.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
$script:buttonAddGroup.ForeColor = [System.Drawing.Color]::White
$script:buttonAddGroup.FlatStyle = "Flat"
$script:buttonAddGroup.Add_Click({ CreateNewGroup })  # 来自GroupOperations.ps1

$script:buttonAddToGroup = New-Object System.Windows.Forms.Button
$script:buttonAddToGroup.Text = "加入组"
$script:buttonAddToGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonAddToGroup.Width = 70
$script:buttonAddToGroup.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
$script:buttonAddToGroup.ForeColor = [System.Drawing.Color]::White
$script:buttonAddToGroup.FlatStyle = "Flat"
$script:buttonAddToGroup.Add_Click({ AddUserToGroup })  # 来自GroupOperations.ps1

$script:buttonModifyGroup = New-Object System.Windows.Forms.Button
$script:buttonModifyGroup.Text = "修改组"
$script:buttonModifyGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonModifyGroup.Width = 70
$script:buttonModifyGroup.BackColor = [System.Drawing.Color]::FromArgb(255, 140, 0)
$script:buttonModifyGroup.ForeColor = [System.Drawing.Color]::White
$script:buttonModifyGroup.FlatStyle = "Flat"
$script:buttonModifyGroup.Add_Click({ ModifyGroup })  # 来自GroupOperations.ps1

$script:buttonDeleteGroup = New-Object System.Windows.Forms.Button
$script:buttonDeleteGroup.Text = "删除组"
$script:buttonDeleteGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonDeleteGroup.Width = 70
$script:buttonDeleteGroup.BackColor = [System.Drawing.Color]::FromArgb(220, 20, 60)
$script:buttonDeleteGroup.ForeColor = [System.Drawing.Color]::White
$script:buttonDeleteGroup.FlatStyle = "Flat"
$script:buttonDeleteGroup.Add_Click({ DeleteGroup })  # 来自GroupOperations.ps1

# 从组移除用户按钮
$script:buttonRemoveFromGroup = New-Object System.Windows.Forms.Button
$script:buttonRemoveFromGroup.Text = "移除组账号"
$script:buttonRemoveFromGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0) 
$script:buttonRemoveFromGroup.Width = 80 
$script:buttonRemoveFromGroup.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 70)
$script:buttonRemoveFromGroup.ForeColor = [System.Drawing.Color]::White
$script:buttonRemoveFromGroup.FlatStyle = "Flat"
$script:buttonRemoveFromGroup.Add_Click({ RemoveUserFromGroup })  # 来自GroupOperations.ps1

$script:groupButtonsPanel.Controls.Add($script:buttonAddGroup)
$script:groupButtonsPanel.Controls.Add($script:buttonAddToGroup)
$script:groupButtonsPanel.Controls.Add($script:buttonModifyGroup)
$script:groupButtonsPanel.Controls.Add($script:buttonDeleteGroup)
$script:groupButtonsPanel.Controls.Add($script:buttonRemoveFromGroup)
$script:groupOpTable.Controls.Add($script:groupButtonsPanel, 1, 3)

# 组选择变化事件（保持原有逻辑）
$script:groupDataGridView.Add_SelectionChanged({
    if ($script:groupDataGridView.SelectedRows.Count -gt 0) {
        $group = $script:groupDataGridView.SelectedRows[0].DataBoundItem
        $script:textGroupName.Text = $group.Name
        $script:textGroupSamAccount.Text = $group.SamAccountName
        $script:textGroupDescription.Text = $group.Description
        $script:originalGroupSamAccount = $group.SamAccountName
    }
})

$script:groupOperationPanel.Controls.Add($script:groupOpTable)
$script:groupManagementPanel.Controls.Add($script:groupOperationPanel, 1, 0)
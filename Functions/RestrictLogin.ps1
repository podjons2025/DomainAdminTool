<# 
核心函数：弹出限制登录计算机窗口
#>

function ShowRestrictLoginForm {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 1. 检查是否选中用户（支持多用户）
    $selectedUsers = $script:userDataGridView.SelectedRows
    if (-not $selectedUsers -or $selectedUsers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "请先在用户列表中选中1个或多个用户", 
            "提示", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
	
	
    # 超级管理员身份校验
    try {
        $adminUsers = @()  # 存储检测到的域管理员用户
        # 遍历每个选中用户，查询其是否属于Domain Admins组
        foreach ($row in $selectedUsers) {
            $samAccountName = $row.Cells["SamAccountName"].Value
            $displayName = $row.Cells["DisplayName"].Value
            $displayName = if ([string]::IsNullOrEmpty($displayName)) { $samAccountName } else { $displayName }

            # 远程查询用户所属的所有安全组（含嵌套组）
            $userGroups = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($sam)
                Import-Module ActiveDirectory -ErrorAction Stop
                # 获取用户的所有安全组，筛选组名包含"Domain Admins"的组
                Get-ADPrincipalGroupMembership -Identity $sam -ErrorAction Stop | 
                    Where-Object { $_.GroupCategory -eq "Security" -and $_.Name -eq "Domain Admins" } | 
                    Select-Object -ExpandProperty Name
            } -ArgumentList $samAccountName -ErrorAction Stop

            # 若存在Domain Admins组，记录该用户
            if ($userGroups -contains "Domain Admins") {
                $adminUsers += "$displayName（账号：$samAccountName）"
            }
        }

        # 若检测到域管理员，弹出告警并终止流程
        if ($adminUsers.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "警告：选中的用户包含超级域控管理员，禁止限制其登录权限！`n`n涉及用户：`n$($adminUsers -join "`n")`n`n请重新选择非管理员用户操作。", 
                "高危操作拦截", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error  # 用错误图标强化警示
            )
            return  # 终止函数，不继续创建窗口
        }
    }
    catch {
        # 处理查询AD时的异常（如权限不足、网络问题）
        [System.Windows.Forms.MessageBox]::Show(
            "管理员身份校验失败：$($_.Exception.Message)", 
            "错误", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # 2. 创建新窗口
    $restrictForm = New-Object System.Windows.Forms.Form
    $restrictForm.Text = "限制用户登录计算机(支持账号多选)"
    $restrictForm.Size = New-Object System.Drawing.Size(800, 550)  # 高度增加50适应搜索框
    $restrictForm.StartPosition = "CenterParent"
    $restrictForm.FormBorderStyle = "FixedDialog"
    $restrictForm.MaximizeBox = $false
    $restrictForm.MinimizeBox = $false

    # 创建ToolTip组件用于显示按钮提示
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 5000  # 提示显示时间（毫秒）
    $toolTip.InitialDelay = 1000  # 鼠标悬停后延迟显示时间
    $toolTip.ReshowDelay = 500    # 连续显示提示的延迟时间
    $toolTip.ShowAlways = $true   # 即使窗口不在焦点也显示

    # 存储显示文本到原始主机名的映射（关键）
    $script:computerNameMap = @{}
    # 存储原始列表数据（用于搜索过滤）
    $script:allowedOriginalItems = @()
    $script:allComputersOriginalItems = @()

    # 3. 窗口布局：使用TableLayoutPanel排版
    $mainTable = New-Object System.Windows.Forms.TableLayoutPanel
    $mainTable.Dock = "Fill"
    $mainTable.Padding = 10
    $mainTable.ColumnCount = 5  # 左列表(2列) + 按钮列(1列) + 右列表(2列)
    $mainTable.RowCount = 2     # 列表区(1行) + 按钮区(1行)
    $mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
    $mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 10)))  # 间距
    $mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 65))) # 按钮列
    $mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 10)))  # 间距
    $mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 85)))
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))

    #4. 左侧列表：允许登录的计算机（增加搜索功能）
    # 重构为TableLayoutPanel用于布局标题、搜索框、列表、数量统计
    $leftPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $leftPanel.Dock = "Fill"
    $leftPanel.RowCount = 4  # 标题行、搜索行、列表行、数量统计行
    $leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 25)))  # 标题高度
    $leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))  # 搜索框高度
    $leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # 列表占比
    $leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))  # 统计高度

    # 标题标签
    $lblAllowed = New-Object System.Windows.Forms.Label
    $lblAllowed.Text = "允许登录的计算机（选中用户共用）"
    $lblAllowed.Dock = "Fill"
    $lblAllowed.Font = New-Object System.Drawing.Font($lblAllowed.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $leftPanel.Controls.Add($lblAllowed, 0, 0)

    # 搜索框区域
    $searchAllowedPanel = New-Object System.Windows.Forms.Panel
    $searchAllowedPanel.Dock = "Fill"
    $searchAllowedPanel.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)

    $lblSearchAllowed = New-Object System.Windows.Forms.Label
    $lblSearchAllowed.Text = "搜索："
    $lblSearchAllowed.Location = New-Object System.Drawing.Point(0, 5)
    $lblSearchAllowed.AutoSize = $true

    $txtSearchAllowed = New-Object System.Windows.Forms.TextBox
    $txtSearchAllowed.Dock = "Fill"
    $txtSearchAllowed.Margin = New-Object System.Windows.Forms.Padding(35, 0, 0, 0)  # 留出标签位置
    $toolTip.SetToolTip($txtSearchAllowed, "支持模糊搜索，不区分大小写")
    # 模拟占位文本
    $allowedPlaceholder = "输入主机名或IP搜索..."
    $txtSearchAllowed.Text = $allowedPlaceholder
    $txtSearchAllowed.ForeColor = [System.Drawing.Color]::Gray  # 占位文本灰色

    # 绑定焦点事件实现占位文本效果
    $txtSearchAllowed.Add_GotFocus({
        if ($this.Text -eq $allowedPlaceholder) {
            $this.Text = ""
            $this.ForeColor = [System.Drawing.Color]::Black  # 输入文本黑色
        }
    })
    $txtSearchAllowed.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($this.Text)) {
            $this.Text = $allowedPlaceholder
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })

    $searchAllowedPanel.Controls.Add($txtSearchAllowed)
    $searchAllowedPanel.Controls.Add($lblSearchAllowed)
    $leftPanel.Controls.Add($searchAllowedPanel, 0, 1)

    # 列表框
    $lstAllowed = New-Object System.Windows.Forms.ListBox
    $lstAllowed.Dock = "Fill"
    $lstAllowed.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended  # 支持多选
    $lstAllowed.IntegralHeight = $false  # 禁用自动高度（适应面板）
    $lstAllowed.ScrollAlwaysVisible = $true  # 始终显示滚动条
    $leftPanel.Controls.Add($lstAllowed, 0, 2)

    # 数量统计标签（底部）
    $lblAllowedCount = New-Object System.Windows.Forms.Label
    $lblAllowedCount.Text = "允许登录计算机共 0 台"
    $lblAllowedCount.Dock = "Fill"
    $lblAllowedCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblAllowedCount.Font = New-Object System.Drawing.Font($lblAllowedCount.Font.FontFamily, 9)
    $leftPanel.Controls.Add($lblAllowedCount, 0, 3)

    $mainTable.Controls.Add($leftPanel, 0, 0)  # 左列表放在第0列第0行

    # 5. 中间按钮：列表移动控制（调整位置至垂直居中）
    # 使用FlowLayoutPanel控制按钮布局
    $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnPanel.Dock = "Fill"
    # 增大顶部内边距使按钮下移至中间（原30调整为120，根据窗口高度计算）
    $btnPanel.Padding = New-Object System.Windows.Forms.Padding(0, 50, 0, 0)  # 顶部留白增加，使按钮居中
    $btnPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown  # 按钮垂直排列
    $btnPanel.WrapContents = $false  # 不自动换行
    $btnPanel.AutoSize = $false  # 不自动调整大小

    # 按钮布局：垂直排列，保持间距
    $btnAddAll = New-Object System.Windows.Forms.Button
    $btnAddAll.Text = "<<"
    $btnAddAll.Width = 60  # 固定按钮宽度
    $btnAddAll.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)  # 上下各8像素间距
    $toolTip.SetToolTip($btnAddAll, "添加所有域内计算机到允许列表")

    $btnAddSelected = New-Object System.Windows.Forms.Button
    $btnAddSelected.Text = "<"
    $btnAddSelected.Width = 60
    $btnAddSelected.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)
    $toolTip.SetToolTip($btnAddSelected, "添加选中的计算机到允许列表")

    $btnRemoveSelected = New-Object System.Windows.Forms.Button
    $btnRemoveSelected.Text = ">"
    $btnRemoveSelected.Width = 60
    $btnRemoveSelected.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)
    $toolTip.SetToolTip($btnRemoveSelected, "从允许列表移除选中的计算机")

    $btnRemoveAll = New-Object System.Windows.Forms.Button
    $btnRemoveAll.Text = ">>"
    $btnRemoveAll.Width = 60
    $btnRemoveAll.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)
    $toolTip.SetToolTip($btnRemoveAll, "清空允许列表")

    # 按顺序添加按钮
    $btnPanel.Controls.Add($btnAddAll)
    $btnPanel.Controls.Add($btnAddSelected)
    $btnPanel.Controls.Add($btnRemoveSelected)
    $btnPanel.Controls.Add($btnRemoveAll)
    $mainTable.Controls.Add($btnPanel, 2, 0)  # 按钮放在第2列第0行

    # 6. 右侧列表：域内所有计算机
    # 重构为TableLayoutPanel用于布局标题、搜索框、列表、数量统计
    $rightPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $rightPanel.Dock = "Fill"
    $rightPanel.RowCount = 4  # 标题行、搜索行、列表行、数量统计行
    $rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 25)))  # 标题高度
    $rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))  # 搜索框高度
    $rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # 列表占比
    $rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))  # 统计高度

    # 标题标签
    $lblAllComputers = New-Object System.Windows.Forms.Label
    $lblAllComputers.Text = "域内所有计算机（从域控检索）"
    $lblAllComputers.Dock = "Fill"
    $lblAllComputers.Font = New-Object System.Drawing.Font($lblAllComputers.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $rightPanel.Controls.Add($lblAllComputers, 0, 0)

    # 搜索框区域
    $searchAllPanel = New-Object System.Windows.Forms.Panel
    $searchAllPanel.Dock = "Fill"
    $searchAllPanel.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)

    $lblSearchAll = New-Object System.Windows.Forms.Label
    $lblSearchAll.Text = "搜索："
    $lblSearchAll.Location = New-Object System.Drawing.Point(0, 5)
    $lblSearchAll.AutoSize = $true

    $txtSearchAll = New-Object System.Windows.Forms.TextBox
    $txtSearchAll.Dock = "Fill"
    $txtSearchAll.Margin = New-Object System.Windows.Forms.Padding(35, 0, 0, 0)  # 留出标签位置
    $toolTip.SetToolTip($txtSearchAll, "支持模糊搜索，不区分大小写")
    # 模拟占位文本
    $allPlaceholder = "输入主机名或IP搜索..."
    $txtSearchAll.Text = $allPlaceholder
    $txtSearchAll.ForeColor = [System.Drawing.Color]::Gray  # 占位文本灰色

    # 绑定焦点事件实现占位文本效果
    $txtSearchAll.Add_GotFocus({
        if ($this.Text -eq $allPlaceholder) {
            $this.Text = ""
            $this.ForeColor = [System.Drawing.Color]::Black  # 输入文本黑色
        }
    })
    $txtSearchAll.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($this.Text)) {
            $this.Text = $allPlaceholder
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })

    $searchAllPanel.Controls.Add($txtSearchAll)
    $searchAllPanel.Controls.Add($lblSearchAll)
    $rightPanel.Controls.Add($searchAllPanel, 0, 1)

    # 列表框
    $lstAllComputers = New-Object System.Windows.Forms.ListBox
    $lstAllComputers.Dock = "Fill"
    $lstAllComputers.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
    $lstAllComputers.IntegralHeight = $false
    $lstAllComputers.ScrollAlwaysVisible = $true
    $rightPanel.Controls.Add($lstAllComputers, 0, 2)

    # 数量统计标签（底部）
    $lblAllComputersCount = New-Object System.Windows.Forms.Label
    $lblAllComputersCount.Text = "列表计算机共 0 台"
    $lblAllComputersCount.Dock = "Fill"
    $lblAllComputersCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblAllComputersCount.Font = New-Object System.Drawing.Font($lblAllComputersCount.Font.FontFamily, 9)
    $rightPanel.Controls.Add($lblAllComputersCount, 0, 3)

    $mainTable.Controls.Add($rightPanel, 4, 0)  # 右列表放在第4列第0行

    # 7. 底部按钮：保存/取消
    $bottomBtnPanel = New-Object System.Windows.Forms.Panel
    $bottomBtnPanel.Dock = "Fill"
    $bottomBtnPanel.Padding = 5

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "保存限制设置"
    $btnSave.Location = New-Object System.Drawing.Point(250, 10)
    $btnSave.Width = 120
    $btnSave.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $btnSave.ForeColor = [System.Drawing.Color]::White
    $btnSave.FlatStyle = "Flat"

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "取消"
    $btnCancel.Location = New-Object System.Drawing.Point(400, 10)
    $btnCancel.Width = 120
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(169, 169, 169)
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = "Flat"

    $bottomBtnPanel.Controls.Add($btnSave)
    $bottomBtnPanel.Controls.Add($btnCancel)
    $mainTable.Controls.Add($bottomBtnPanel, 0, 1)  # 底部按钮放在第0列第1行
    $mainTable.SetColumnSpan($bottomBtnPanel, 5)  # 跨5列（占满底部宽度）

    # 8. 数量统计更新函数
    function UpdateCounts {
        $lblAllowedCount.Text = "允许登录计算机共 $($lstAllowed.Items.Count) 台"
        $lblAllComputersCount.Text = "列表计算机共 $($lstAllComputers.Items.Count) 台"
    }

    # 9. 搜索过滤函数（处理占位文本情况）
    function FilterList($sourceList, $targetList, $originalItems, $searchText, $placeholder) {
        # 如果是占位文本，视为空搜索
        $actualSearchText = if ($searchText -eq $placeholder) { "" } else { $searchText }
        
        $targetList.Items.Clear()
        if ([string]::IsNullOrWhiteSpace($actualSearchText)) {
            # 搜索为空时显示所有原始项
            $originalItems | ForEach-Object { $targetList.Items.Add($_) | Out-Null }
        }
        else {
            # 模糊匹配（不区分大小写）
            $lowerSearch = $actualSearchText.ToLower()
            $originalItems | Where-Object { $_.ToLower().Contains($lowerSearch) } | ForEach-Object {
                $targetList.Items.Add($_) | Out-Null
            }
        }
    }

    # 10. 加载域内计算机
    function LoadDomainComputers {
        try {
            # 清空映射表和原始列表
            $script:computerNameMap.Clear()
            $script:allowedOriginalItems = @()
            $script:allComputersOriginalItems = @()

            # 远程从域控获取所有计算机（包含IP地址）
            $allComputers = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                Import-Module ActiveDirectory -ErrorAction Stop
                # 获取所有启用的计算机，包含名称和IP地址
                Get-ADComputer -Filter { Enabled -eq $true } -Properties Name, IPv4Address | 
                    Select-Object Name, IPv4Address | 
                    Sort-Object Name  # 按名称排序
            } -ErrorAction Stop

            # 处理计算机数据，生成"主机名--（IP）"格式的显示文本
            $displayTexts = @()
            foreach ($comp in $allComputers) {
                $hostName = $comp.Name
                $ipAddress = $comp.IPv4Address
                
                # 处理无IP地址的情况
                $displayIp = if ($ipAddress) { $ipAddress } else { "无IP" }
                $displayText = "$hostName -- ($displayIp)"
                
                # 添加到显示列表和映射表
                $displayTexts += $displayText
                $script:computerNameMap[$displayText] = $hostName  # 映射显示文本到原始主机名
            }

            # 保存右侧列表原始数据并初始化显示
            $script:allComputersOriginalItems = $displayTexts
            $lstAllComputers.Items.Clear()
            $displayTexts | ForEach-Object { $lstAllComputers.Items.Add($_) | Out-Null }

            # 加载选中用户的【现有允许登录计算机】
            $firstUserWorkstations = $null
            $hasDifferentSettings = $false
            foreach ($row in $selectedUsers) {
                $samAccountName = $row.Cells["SamAccountName"].Value
                $userWorkstations = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                    param($sam)
                    Import-Module ActiveDirectory -ErrorAction Stop
                    Get-ADUser -Identity $sam -Properties userWorkstations -ErrorAction Stop | 
                        Select-Object -ExpandProperty userWorkstations
                } -ArgumentList $samAccountName -ErrorAction Stop

                # 首次加载时记录第一个用户的设置
                if (-not $firstUserWorkstations) {
                    $firstUserWorkstations = $userWorkstations
                }
                # 检查多用户设置是否一致
                elseif ($userWorkstations -ne $firstUserWorkstations) {
                    $hasDifferentSettings = $true
                }
            }

            # 处理左侧列表数据
            $allowedDisplayTexts = @()
            if ($hasDifferentSettings) {
                [System.Windows.Forms.MessageBox]::Show("选中的用户现有登录限制设置不一致，将统一按本次选择重新设置", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            elseif ($firstUserWorkstations) {
                # 拆分现有计算机名（按逗号分割，去空）
                $allowedHostNames = $firstUserWorkstations -split "," | 
                    ForEach-Object { $_.Trim() } | 
                    Where-Object { $_ -ne "" }

                # 转换为主机名--（IP）格式显示
                foreach ($hostName in $allowedHostNames) {
                    # 查找对应的显示文本
                    $mappedText = $script:computerNameMap.GetEnumerator() | 
                        Where-Object { $_.Value -eq $hostName } | 
                        Select-Object -ExpandProperty Key -First 1

                    if ($mappedText) {
                        $allowedDisplayTexts += $mappedText
                    }
                    else {
                        # 处理已不存在的计算机（在域中找不到）
                        $customText = "$hostName--（已不存在）"
                        $allowedDisplayTexts += $customText
                        # 添加到映射表，确保保存时能正确提取原始主机名
                        $script:computerNameMap[$customText] = $hostName
                    }
                }
            }

            # 保存左侧列表原始数据并初始化显示
            $script:allowedOriginalItems = $allowedDisplayTexts
            $lstAllowed.Items.Clear()
            $allowedDisplayTexts | ForEach-Object { $lstAllowed.Items.Add($_) | Out-Null }

            # 加载完成后更新数量统计
            UpdateCounts

        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("检索计算机失败：$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $restrictForm.Close()
        }
    }

    # 11. 绑定搜索框事件（传递占位文本参数）
    $txtSearchAllowed.Add_TextChanged({
        FilterList -sourceList $lstAllowed `
                   -targetList $lstAllowed `
                   -originalItems $script:allowedOriginalItems `
                   -searchText $txtSearchAllowed.Text `
                   -placeholder $allowedPlaceholder
        UpdateCounts
    })

    $txtSearchAll.Add_TextChanged({
        FilterList -sourceList $lstAllComputers `
                   -targetList $lstAllComputers `
                   -originalItems $script:allComputersOriginalItems `
                   -searchText $txtSearchAll.Text `
                   -placeholder $allPlaceholder
        UpdateCounts
    })

    # 窗口加载时执行计算机检索
    $restrictForm.Add_Shown({ LoadDomainComputers })

# 12. 列表移动按钮逻辑
    # 全部添加到允许列表
    $btnAddAll.Add_Click({
        # 基于原始数据操作，避免搜索过滤影响
        foreach ($item in $script:allComputersOriginalItems) {
            if (-not $script:allowedOriginalItems.Contains($item)) {
                $script:allowedOriginalItems += $item
            }
        }
        # 刷新左侧列表（保持当前搜索状态）
        FilterList -sourceList $lstAllowed `
                   -targetList $lstAllowed `
                   -originalItems $script:allowedOriginalItems `
                   -searchText $txtSearchAllowed.Text `
                   -placeholder $allowedPlaceholder
        UpdateCounts  # 更新数量统计
    })

    # 选中添加到允许列表
    $btnAddSelected.Add_Click({
        $selectedItems = @($lstAllComputers.SelectedItems)
        foreach ($item in $selectedItems) {
            if (-not $script:allowedOriginalItems.Contains($item)) {
                $script:allowedOriginalItems += $item
            }
        }
        # 刷新左侧列表（保持当前搜索状态）
        FilterList -sourceList $lstAllowed `
                   -targetList $lstAllowed `
                   -originalItems $script:allowedOriginalItems `
                   -searchText $txtSearchAllowed.Text `
                   -placeholder $allowedPlaceholder
        $lstAllComputers.ClearSelected()  # 取消右侧选中
        UpdateCounts  # 更新数量统计
    })

    # 选中从允许列表移除
    $btnRemoveSelected.Add_Click({
        $selectedItems = @($lstAllowed.SelectedItems)
        foreach ($item in $selectedItems) {
            $script:allowedOriginalItems = $script:allowedOriginalItems | Where-Object { $_ -ne $item }
        }
        # 刷新左侧列表（保持当前搜索状态）
        FilterList -sourceList $lstAllowed `
                   -targetList $lstAllowed `
                   -originalItems $script:allowedOriginalItems `
                   -searchText $txtSearchAllowed.Text `
                   -placeholder $allowedPlaceholder
        $lstAllowed.ClearSelected()  # 取消左侧选中
        UpdateCounts  # 更新数量统计
    })

    # 清空允许列表
    $btnRemoveAll.Add_Click({
        if ([System.Windows.Forms.MessageBox]::Show("确定要清空所有允许登录的计算机吗？", "确认", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq "Yes") {
            $script:allowedOriginalItems = @()
            # 刷新左侧列表（保持当前搜索状态）
            FilterList -sourceList $lstAllowed `
                       -targetList $lstAllowed `
                       -originalItems $script:allowedOriginalItems `
                       -searchText $txtSearchAllowed.Text `
                       -placeholder $allowedPlaceholder
            UpdateCounts  # 更新数量统计
        }
    })
    

# 13. 保存限制设置（提取原始主机名）
    $btnSave.Add_Click({
        # 1. 从原始允许列表提取所有主机名（不受搜索过滤影响）
        $allowedHostNames = @()
        foreach ($displayItem in $script:allowedOriginalItems) {
            # 从映射表获取原始主机名
            if ($script:computerNameMap.ContainsKey($displayItem)) {
                $allowedHostNames += $script:computerNameMap[$displayItem]
            }
            else {
                # 从显示文本中提取主机名（按格式拆分）
                $hostNamePart = $displayItem -split "--", 2 | Select-Object -First 1
                $allowedHostNames += $hostNamePart.Trim()
            }
        }
        $allowedComputers = $allowedHostNames -join ","
        $isClearRestriction = [string]::IsNullOrEmpty($allowedComputers)  # 判断是否为“撤离限制”

        # 2. 确认保存
        $confirmMsg = if ($isClearRestriction) {
            "将为 $($selectedUsers.Count) 个用户【撤离登录限制】（允许登录所有计算机）`n`n确定保存吗？"
        } else {
            "将为 $($selectedUsers.Count) 个用户设置允许登录的计算机：`n$allowedComputers`n`n确定保存吗？"
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            $confirmMsg, 
            "确认保存", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne "Yes") { return }

        try {
            # 记录成功/失败的用户详情
            $successUsers = @()  # 格式："显示名（账号名）"
            $failUsers = @()     # 格式："显示名（账号名）- 失败原因"

            # 遍历每个选中用户，批量设置
            foreach ($row in $selectedUsers) {
                $displayName = $row.Cells["DisplayName"].Value
                $samAccountName = $row.Cells["SamAccountName"].Value
                $displayName = if ([string]::IsNullOrEmpty($displayName)) { "未设置显示名" } else { $displayName }
                $userLabel = "$displayName（$samAccountName）"

                try {
                    Invoke-Command -Session $script:remoteSession -ScriptBlock {
                        param($sam, $workstations, $isClear)
                        Import-Module ActiveDirectory -ErrorAction Stop

                        if ($isClear) {
                            # 撤离限制：清除userWorkstations属性
                            Set-ADUser -Identity $sam -Clear userWorkstations -ErrorAction Stop
                        } else {
                            # 设置限制：更新userWorkstations属性
                            Set-ADUser -Identity $sam -Replace @{userWorkstations = $workstations} -ErrorAction Stop
                        }
                    } -ArgumentList $samAccountName, $allowedComputers, $isClearRestriction -ErrorAction Stop

                    $successUsers += $userLabel
                }
                catch {
                    $failReason = $_.Exception.Message
                    $failUsers += "$userLabel - $failReason"
                }
            }

            # 3. 生成结果提示
            $resultTitle = if ($failUsers.Count -eq 0) { "成功" } else { "结果" }
            $resultIcon = if ($failUsers.Count -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
            
            $resultMsg = "设置完成！`n"
            $resultMsg += "`n成功用户（共 $($successUsers.Count) 个）："
            if ($successUsers.Count -eq 0) {
                $resultMsg += "`n无"
            } else {
                $resultMsg += "`n$($successUsers -join "`n")"
            }
            $resultMsg += "`n`n失败用户（共 $($failUsers.Count) 个）："
            if ($failUsers.Count -eq 0) {
                $resultMsg += "`n无"
            } else {
                $resultMsg += "`n$($failUsers -join "`n")"
            }

            [System.Windows.Forms.MessageBox]::Show($resultMsg, $resultTitle, [System.Windows.Forms.MessageBoxButtons]::OK, $resultIcon)

            # 全部成功时自动关闭窗口
            if ($failUsers.Count -eq 0) {
                $restrictForm.Close()
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("保存流程异常：$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    


    # 14. 取消按钮逻辑
    $btnCancel.Add_Click({
        $restrictForm.Close()
    })

    # 15. 窗口显示
    $restrictForm.Controls.Add($mainTable)
    $restrictForm.ShowDialog() | Out-Null

    # 清理脚本级变量
    Remove-Variable -Name computerNameMap -Scope Script -ErrorAction SilentlyContinue
    Remove-Variable -Name allowedOriginalItems -Scope Script -ErrorAction SilentlyContinue
    Remove-Variable -Name allComputersOriginalItems -Scope Script -ErrorAction SilentlyContinue
}
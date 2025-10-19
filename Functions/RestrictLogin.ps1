<# 
核心函数：弹出批量限制登录计算机窗口（双列表版）
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
	
	
    # ---------------------- 新增：超级管理员身份校验 ----------------------
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
    # ---------------------- 新增结束 ----------------------	

    # 2. 创建新窗口
    $restrictForm = New-Object System.Windows.Forms.Form
    $restrictForm.Text = "限制用户登录计算机(支持账号多选)"
    $restrictForm.Size = New-Object System.Drawing.Size(800, 500)
    $restrictForm.StartPosition = "CenterParent"
    $restrictForm.FormBorderStyle = "FixedDialog"
    $restrictForm.MaximizeBox = $false
    $restrictForm.MinimizeBox = $false

    # 新增：创建ToolTip组件用于显示按钮提示
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 5000  # 提示显示时间（毫秒）
    $toolTip.InitialDelay = 1000  # 鼠标悬停后延迟显示时间
    $toolTip.ReshowDelay = 500    # 连续显示提示的延迟时间
    $toolTip.ShowAlways = $true   # 即使窗口不在焦点也显示

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

    # ---------------------- 4. 左侧列表：允许登录的计算机 ----------------------
    $leftPanel = New-Object System.Windows.Forms.Panel
    $leftPanel.Dock = "Fill"

    $lblAllowed = New-Object System.Windows.Forms.Label
    $lblAllowed.Text = "允许登录的计算机（选中用户共用）"
    $lblAllowed.Dock = "Top"
    $lblAllowed.Font = New-Object System.Drawing.Font($lblAllowed.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)

    $lstAllowed = New-Object System.Windows.Forms.ListBox
    $lstAllowed.Dock = "Fill"
    $lstAllowed.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended  # 支持多选
    $lstAllowed.IntegralHeight = $false  # 禁用自动高度（适应面板）
    $lstAllowed.ScrollAlwaysVisible = $true  # 始终显示滚动条

    $leftPanel.Controls.Add($lstAllowed)
    $leftPanel.Controls.Add($lblAllowed)  # 标签在列表上方
    $mainTable.Controls.Add($leftPanel, 0, 0)  # 左列表放在第0列第0行

    # ---------------------- 5. 中间按钮：列表移动控制----------------------
    $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnPanel.Dock = "Fill"
    $btnPanel.Padding = New-Object System.Windows.Forms.Padding(0, 30, 0, 0)  # 顶部留30像素空白，使按钮整体下移
    $btnPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown  # 按钮垂直排列
    $btnPanel.WrapContents = $false  # 不自动换行
    $btnPanel.AutoSize = $false  # 不自动调整大小

    # 按钮布局：垂直排列，增加间距
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

    # 按顺序添加按钮（FlowDirection=TopDown，先加的在上面）
    $btnPanel.Controls.Add($btnAddAll)
    $btnPanel.Controls.Add($btnAddSelected)
    $btnPanel.Controls.Add($btnRemoveSelected)
    $btnPanel.Controls.Add($btnRemoveAll)
    $mainTable.Controls.Add($btnPanel, 2, 0)  # 按钮放在第2列第0行

    # ---------------------- 6. 右侧列表：域内所有计算机 ----------------------
    $rightPanel = New-Object System.Windows.Forms.Panel
    $rightPanel.Dock = "Fill"

    $lblAllComputers = New-Object System.Windows.Forms.Label
    $lblAllComputers.Text = "域内所有计算机（从域控检索）"
    $lblAllComputers.Dock = "Top"
    $lblAllComputers.Font = New-Object System.Drawing.Font($lblAllComputers.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)

    $lstAllComputers = New-Object System.Windows.Forms.ListBox
    $lstAllComputers.Dock = "Fill"
    $lstAllComputers.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
    $lstAllComputers.IntegralHeight = $false
    $lstAllComputers.ScrollAlwaysVisible = $true

    $rightPanel.Controls.Add($lstAllComputers)
    $rightPanel.Controls.Add($lblAllComputers)
    $mainTable.Controls.Add($rightPanel, 4, 0)  # 右列表放在第4列第0行

    # ---------------------- 7. 底部按钮：保存/取消 ----------------------
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

    # ---------------------- 8. 加载域内计算机（核心逻辑） ----------------------
    function LoadDomainComputers {
        try {
            # 远程从域控获取所有计算机（只取计算机名，去重）
            $allComputers = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                Import-Module ActiveDirectory -ErrorAction Stop
                # 获取所有启用的计算机，排除重复（按Name去重）
                Get-ADComputer -Filter { Enabled -eq $true } -Properties Name | 
                    Select-Object -ExpandProperty Name -Unique | 
                    Sort-Object  # 按名称排序
            } -ErrorAction Stop

            # 填充右侧列表
            $lstAllComputers.Items.Clear()
            $allComputers | ForEach-Object { $lstAllComputers.Items.Add($_) | Out-Null }

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

            # 填充左侧列表
            $lstAllowed.Items.Clear()
            if ($hasDifferentSettings) {
                [System.Windows.Forms.MessageBox]::Show("选中的用户现有登录限制设置不一致，将统一按本次选择重新设置", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            elseif ($firstUserWorkstations) {
                # 拆分现有计算机名（按逗号分割，去空）
                $firstUserWorkstations -split "," | 
                    ForEach-Object { $_.Trim() } | 
                    Where-Object { $_ -ne "" } | 
                    ForEach-Object { $lstAllowed.Items.Add($_) | Out-Null }
            }

        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("检索计算机失败：$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $restrictForm.Close()
        }
    }

    # 窗口加载时执行计算机检索
    $restrictForm.Add_Shown({ LoadDomainComputers })

# ---------------------- 9. 列表移动按钮逻辑 ----------------------
    # 全部添加到允许列表
    $btnAddAll.Add_Click({
        $allItems = [array]$lstAllComputers.Items
        foreach ($item in $allItems) {
            if (-not $lstAllowed.Items.Contains($item)) {
                $lstAllowed.Items.Add($item) | Out-Null
            }
        }
    })

    # 选中添加到允许列表
    $btnAddSelected.Add_Click({
        $selectedItems = @($lstAllComputers.SelectedItems)
        foreach ($item in $selectedItems) {
            if (-not $lstAllowed.Items.Contains($item)) {
                $lstAllowed.Items.Add($item) | Out-Null
            }
        }
        # 取消右侧选中（提升体验）
        $lstAllComputers.ClearSelected()
    })

    # 选中从允许列表移除
    $btnRemoveSelected.Add_Click({
        $selectedItems = [array]$lstAllowed.SelectedItems
        foreach ($item in $selectedItems) {
            $lstAllowed.Items.Remove($item) | Out-Null
        }
        # 取消左侧选中
        $lstAllowed.ClearSelected()
    })

    # 清空允许列表
    $btnRemoveAll.Add_Click({
        if ([System.Windows.Forms.MessageBox]::Show("确定要清空所有允许登录的计算机吗？", "确认", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq "Yes") {
            $lstAllowed.Items.Clear()
        }
    })
    

# ---------------------- 10. 保存限制设置（核心逻辑） ----------------------
    $btnSave.Add_Click({
        # 1. 处理允许登录计算机列表（空列表时标记为清空）
        $allowedComputers = ($lstAllowed.Items | ForEach-Object { $_ }) -join ","
        $isClearRestriction = [string]::IsNullOrEmpty($allowedComputers)  # 判断是否为“撤离限制”（清空列表）

        # 2. 确认保存（提示信息优化，增加“清空限制”说明）
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
            # 新增：记录成功/失败的用户详情（含DisplayName和SamAccountName）
            $successUsers = @()  # 格式："显示名（账号名）"
            $failUsers = @()     # 格式："显示名（账号名）- 失败原因"

            # 遍历每个选中用户，批量设置
            foreach ($row in $selectedUsers) {
                # 获取用户的DisplayName和SamAccountName（从选中行的单元格中读取）
                $displayName = $row.Cells["DisplayName"].Value
                $samAccountName = $row.Cells["SamAccountName"].Value
                # 处理字段为空的情况（避免显示异常）
                $displayName = if ([string]::IsNullOrEmpty($displayName)) { "未设置显示名" } else { $displayName }
                $userLabel = "$displayName（$samAccountName）"  # 统一的用户标识格式

                try {
                    Invoke-Command -Session $script:remoteSession -ScriptBlock {
                        param($sam, $workstations, $isClear)
                        Import-Module ActiveDirectory -ErrorAction Stop

                        # 清空限制用-Clear，设置限制用-Replace
                        if ($isClear) {
                            # 撤离限制：清除userWorkstations属性（允许登录所有计算机）
                            Set-ADUser -Identity $sam -Clear userWorkstations -ErrorAction Stop
                        } else {
                            # 设置限制：更新userWorkstations属性（指定允许登录的计算机）
                            Set-ADUser -Identity $sam -Replace @{userWorkstations = $workstations} -ErrorAction Stop
                        }
                    } -ArgumentList $samAccountName, $allowedComputers, $isClearRestriction -ErrorAction Stop

                    # 成功：添加到成功列表
                    $successUsers += $userLabel
                }
                catch {
                    # 失败：添加到失败列表（含原因）
                    $failReason = $_.Exception.Message
                    $failUsers += "$userLabel - $failReason"
                }
            }

            # 3. 生成结果提示（完整显示成功/失败用户详情）
            $resultTitle = if ($failUsers.Count -eq 0) { "成功" } else { "结果" }
            $resultIcon = if ($failUsers.Count -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
            
            # 拼接结果内容
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

            # 显示结果
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
    


    # ---------------------- 11. 取消按钮逻辑 ----------------------
    $btnCancel.Add_Click({
        $restrictForm.Close()
    })

    # ---------------------- 12. 窗口显示 ----------------------
    $restrictForm.Controls.Add($mainTable)
    $restrictForm.ShowDialog() | Out-Null
}
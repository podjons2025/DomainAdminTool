<# 
核心函数：弹出限制登录时间窗口
#>
function ShowRestrictLogonTimeForm {
    # 前置检查：域控连接状态
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 1. 获取域控一周起始日（动态适配不同域控设置）
    try {
        $domainFirstDayOfWeek = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            (Get-WinSystemLocale).DateTimeFormat.FirstDayOfWeek.value__
        } -ErrorAction Stop

        # 动态生成「AD字节索引→星期」映射
        $script:adWeekDays = switch ($domainFirstDayOfWeek) {
            0 { @("周日", "周一", "周二", "周三", "周四", "周五", "周六") }
            1 { @("周一", "周二", "周三", "周四", "周五", "周六", "周日") }
            default { @("周一", "周二", "周三", "周四", "周五", "周六", "周日") }
        }

        # 动态生成「UI选择星期→AD索引」映射
        $script:uiToAdWeekDay = @{}
        for ($i = 0; $i -lt $script:adWeekDays.Count; $i++) {
            $script:uiToAdWeekDay[$script:adWeekDays[$i]] = $i
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("获取域控时间设置失败：$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 2. 检查用户选择
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

    # 3. 创建弹出窗口
    $restrictTimeForm = New-Object System.Windows.Forms.Form
    $restrictTimeForm.Text = "限制用户登录时间(支持账号多选)"
    $restrictTimeForm.Size = New-Object System.Drawing.Size(700, 700)
    $restrictTimeForm.StartPosition = "CenterParent"
    $restrictTimeForm.FormBorderStyle = "FixedDialog"
    $restrictTimeForm.MaximizeBox = $false
    $restrictTimeForm.MinimizeBox = $false

    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 5000
    $toolTip.InitialDelay = 1000
    $toolTip.ReshowDelay = 500
    $toolTip.ShowAlways = $true

    # 4. 窗口布局（4行：标题复选框→工作日全选框→时间段列表→底部按钮）
    $mainTable = New-Object System.Windows.Forms.TableLayoutPanel
    $mainTable.Dock = "Fill"
    $mainTable.Padding = 10
    $mainTable.ColumnCount = 1
    $mainTable.RowCount = 4
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50)))    # 允许所有时间复选框
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))    # 工作日全选复选框
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))   # 时间段列表
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)))    # 底部按钮

    # 5. 「允许所有时间登录」复选框
    $chkAllowAllTime = New-Object System.Windows.Forms.CheckBox
    $chkAllowAllTime.Text = "允许所有时间登录（忽略下方时段选择）"
    $chkAllowAllTime.Font = New-Object System.Drawing.Font($chkAllowAllTime.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $chkAllowAllTime.AutoSize = $false
    $chkAllowAllTime.Width = $restrictTimeForm.ClientSize.Width - 40
    $chkAllowAllTime.Height = 40
    $chkAllowAllTime.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $chkAllowAllTime.Padding = New-Object System.Windows.Forms.Padding(2, 5, 2, 5)
    
    $chkAllowAllTime.Add_CheckedChanged({
        $lstLogonTime.Enabled = -not $chkAllowAllTime.Checked
        # 同步禁用/启用工作日全选框
        $script:weekDayCheckboxes | ForEach-Object { $_.Enabled = -not $chkAllowAllTime.Checked }
    })
    $mainTable.Controls.Add($chkAllowAllTime, 0, 0)

    # 6. 工作日全选复选框（修复了ColumnStyle的百分比计算问题）
    $weekDayTable = New-Object System.Windows.Forms.TableLayoutPanel
    $weekDayTable.Dock = "Fill"
    $weekDayTable.ColumnCount = 7  # 7个工作日各占一列
    
    # 清除默认列样式
    while ($weekDayTable.ColumnStyles.Count -gt 0) {
        $weekDayTable.ColumnStyles.RemoveAt(0)
    }
    
    # 计算每列宽度百分比（使用显式浮点数计算避免错误）
    $columnPercent = 100.0 / 7  # 显式使用浮点数除法
    for ($i = 0; $i -lt 7; $i++) {
        $colStyle = New-Object System.Windows.Forms.ColumnStyle
        $colStyle.SizeType = [System.Windows.Forms.SizeType]::Percent
        $colStyle.Width = $columnPercent  # 使用预计算的百分比值
        $weekDayTable.ColumnStyles.Add($colStyle)
    }

    # 定义UI固定显示的工作日
    $script:weekDaysUI = @("周一", "周二", "周三", "周四", "周五", "周六", "周日")
    $script:weekDayCheckboxes = @()  # 存储所有工作日复选框

    foreach ($day in $script:weekDaysUI) {
        $chkDay = New-Object System.Windows.Forms.CheckBox
        $chkDay.Text = $day
        $chkDay.Font = New-Object System.Drawing.Font($chkDay.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
        $chkDay.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $chkDay.AutoSize = $true
        $toolTip.SetToolTip($chkDay, "全选/取消全选$day所有时段")

        # 全选/取消全选事件
        $chkDay.Add_CheckedChanged({
            param($sender)
            $targetDay = $sender.Text
            $checkedState = $sender.Checked
            # 遍历时间段列表，匹配当前工作日的所有项
            for ($i = 0; $i -lt $lstLogonTime.Items.Count; $i++) {
                $item = $lstLogonTime.Items[$i].ToString()
                if ($item -like "$targetDay(*") {  # 匹配“周一(xx:xx-xx:xx)”格式
                    $lstLogonTime.SetItemChecked($i, $checkedState)
                }
            }
        })

        $weekDayTable.Controls.Add($chkDay)
        $script:weekDayCheckboxes += $chkDay
    }

    $mainTable.Controls.Add($weekDayTable, 0, 1)

    # 7. 登录时段选择列表（与域控一致的时间段格式）
    $lstLogonTime = New-Object System.Windows.Forms.CheckedListBox
    $lstLogonTime.Dock = "Fill"
    $lstLogonTime.IntegralHeight = $false
    $lstLogonTime.ScrollAlwaysVisible = $true
    $lstLogonTime.CheckOnClick = $true
    $lstLogonTime.Font = New-Object System.Drawing.Font($lstLogonTime.Font.FontFamily, 8.5)

    # 生成时间段（格式：周一(0:00-1:00)、周日(23:00-0:00)）
    foreach ($day in $script:weekDaysUI) {
        for ($hour = 0; $hour -lt 24; $hour++) {
            # 计算时间段结束小时（跨天处理：23点结束为0点）
            $endHour = if ($hour -eq 23) { 0 } else { $hour + 1 }
            # 格式化时间段
            $timeSlot = "$day($hour`:00-$endHour`:00)"
            $lstLogonTime.Items.Add($timeSlot, $false)
        }
    }

    $listPanel = New-Object System.Windows.Forms.Panel
    $listPanel.Dock = "Fill"
    $listPanel.Controls.Add($lstLogonTime)
    $mainTable.Controls.Add($listPanel, 0, 2)

    # 8. 底部按钮面板
    $bottomBtnPanel = New-Object System.Windows.Forms.Panel
    $bottomBtnPanel.Dock = "Fill"
    $bottomBtnPanel.Padding = 5

    $btnSaveTime = New-Object System.Windows.Forms.Button
    $btnSaveTime.Text = "保存限制设置"
    $btnSaveTime.Location = New-Object System.Drawing.Point(230, 10)
    $btnSaveTime.Width = 120
    $btnSaveTime.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $btnSaveTime.ForeColor = [System.Drawing.Color]::White
    $btnSaveTime.FlatStyle = "Flat"
    $toolTip.SetToolTip($btnSaveTime, "将所选时段应用到选中用户")

    $btnCancelTime = New-Object System.Windows.Forms.Button
    $btnCancelTime.Text = "取消"
    $btnCancelTime.Location = New-Object System.Drawing.Point(370, 10)
    $btnCancelTime.Width = 120
    $btnCancelTime.BackColor = [System.Drawing.Color]::FromArgb(169, 169, 169)
    $btnCancelTime.ForeColor = [System.Drawing.Color]::White
    $btnCancelTime.FlatStyle = "Flat"

    $bottomBtnPanel.Controls.Add($btnSaveTime)
    $bottomBtnPanel.Controls.Add($btnCancelTime)
    $mainTable.Controls.Add($bottomBtnPanel, 0, 3)

    # ---------------------- 9. 辅助函数1：AD logonHours → UI时间段 ----------------------
    function Convert-LogonHoursToTimeSlots {
        param([byte[]]$logonHours)
        $timeSlots = @()

        # 处理默认值（未设置时允许所有时间）
        if (-not $logonHours -or $logonHours.Length -ne 21) {
            $logonHours = [byte[]]::CreateInstance([byte], 21)
            for ($i = 0; $i -lt 21; $i++) { $logonHours[$i] = 0xFF }
        }

        # 定义本地时间星期→UI星期的映射
        $localWeekDayToUI = @{
            0 = "周日"
            1 = "周一"
            2 = "周二"
            3 = "周三"
            4 = "周四"
            5 = "周五"
            6 = "周六"
        }

        # 遍历AD的21字节（7天×3字节/天）
        for ($byteIndex = 0; $byteIndex -lt 21; $byteIndex++) {
            $currentByte = $logonHours[$byteIndex]
            # 1. 计算当前字节对应的AD星期索引（0-6）
            $adWeekDayIndex = [math]::Floor($byteIndex / 3)
            # 2. 计算当前字节对应的UTC小时段（0-7,8-15,16-23）
            $hourSegment = $byteIndex % 3
            $hourBase = $hourSegment * 8

            # 解析每个Bit对应的UTC小时
            for ($bitIndex = 0; $bitIndex -lt 8; $bitIndex++) {
                $utcHour = $hourBase + $bitIndex
                # 3. 构建UTC时间（使用固定日期确保星期计算正确）
                $utcTime = [DateTime]::new(2023, 1, 1 + $adWeekDayIndex, $utcHour, 0, 0, [DateTimeKind]::Utc)
                # 4. 转换为本地时间
                $localTime = $utcTime.ToLocalTime()
                $localHour = $localTime.Hour
                $endHour = if ($localHour -eq 23) { 0 } else { $localHour + 1 }
                
                # 5. 获取本地时间的星期索引
                $localWeekDayNum = $localTime.DayOfWeek.value__
                # 6. 映射到UI显示的星期名称
                $uiWeekDay = $localWeekDayToUI[$localWeekDayNum]

                # 7. 若当前Bit为1（允许登录），添加时间段
                if (($currentByte -band (1 -shl $bitIndex)) -ne 0) {
                    $timeSlot = "$uiWeekDay($localHour`:00-$endHour`:00)"
                    $timeSlots += $timeSlot
                }
            }
        }
        return $timeSlots
    }

    # ---------------------- 10. 辅助函数2：UI时间段 → AD logonHours（核心修复处） ----------------------
    function Convert-TimeSlotsToLogonHours {
        param([string[]]$selectedTimeSlots)
        $logonHours = New-Object byte[] 21
        
        if (-not $selectedTimeSlots -or $selectedTimeSlots.Count -eq 0) {
            return $logonHours
        }

        # 定义UI星期→本地时间星期索引的映射
        $uiToLocalWeekDay = @{
            "周一" = 1
            "周二" = 2
            "周三" = 3
            "周四" = 4
            "周五" = 5
            "周六" = 6
            "周日" = 0
        }

        foreach ($timeSlot in $selectedTimeSlots) {
            if ($timeSlot -match '^(周一|周二|周三|周四|周五|周六|周日)\((\d+):00-(\d+):00\)$') {
                $uiWeekDay = $matches[1]
                $localHour = [int]$matches[2]
                $endHour = [int]$matches[3]
                
                # 1. 获取本地时间的星期索引
                $localWeekDayNum = $uiToLocalWeekDay[$uiWeekDay]
                
                # 2. 构建本地时间（使用固定日期确保星期计算正确）
                $localTime = [DateTime]::new(2023, 1, 1 + $localWeekDayNum, $localHour, 0, 0, [DateTimeKind]::Local)
                
                # 3. 转换为UTC时间
                $utcTime = $localTime.ToUniversalTime()
                $utcHour = $utcTime.Hour
                
                # 4. 计算AD星期索引（0-6）
                $adWeekDayIndex = $utcTime.DayOfWeek.value__
                
                # 5. 计算字节位置和位索引
                $hourSegment = [math]::Floor($utcHour / 8)
                $byteIndex = $adWeekDayIndex * 3 + $hourSegment
                $bitIndex = $utcHour % 8

                # 6. 验证字节索引合法性
                if ($byteIndex -ge 0 -and $byteIndex -lt 21) {
                    # 7. 置位允许登录
                    $logonHours[$byteIndex] = $logonHours[$byteIndex] -bor (1 -shl $bitIndex)
                }
            }
        }
        return $logonHours
    }

    # 11. 加载用户现有登录时间设置
    function LoadUserLogonTime {
        try {
            $firstUserLogonHours = $null
            $hasDifferentSettings = $false
            $selectedSamAccounts = $selectedUsers | ForEach-Object { $_.Cells["SamAccountName"].Value }

            # 获取第一个用户的登录时间作为基准
            foreach ($row in $selectedUsers) {
                $samAccountName = $row.Cells["SamAccountName"].Value
                $displayName = $row.Cells["DisplayName"].Value
                $displayName = if ([string]::IsNullOrEmpty($displayName)) { $samAccountName } else { $displayName }
                
                # 远程查询用户logonHours
                $userLogonHours = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                    param($sam)
                    Import-Module ActiveDirectory -ErrorAction Stop
                    $user = Get-ADUser -Identity $sam -Properties logonHours -ErrorAction Stop
                    return $user.logonHours
                } -ArgumentList $samAccountName -ErrorAction Stop

                # 对比多用户设置是否一致
                if (-not $firstUserLogonHours) {
                    $firstUserLogonHours = $userLogonHours
                } else {
                    $isSame = $true
                    if (($firstUserLogonHours -eq $null) -xor ($userLogonHours -eq $null)) {
                        $isSame = $false
                    } elseif ($firstUserLogonHours -ne $null) {
                        if ($firstUserLogonHours.Length -ne $userLogonHours.Length) { $isSame = $false }
                        else {
                            for ($i = 0; $i -lt $firstUserLogonHours.Length; $i++) {
                                if ($firstUserLogonHours[$i] -ne $userLogonHours[$i]) { $isSame = $false; break }
                            }
                        }
                    }
                    if (-not $isSame) {
                        $hasDifferentSettings = $true
                    }
                }
            }

            # 转换为新格式的时间段列表
            $allowedTimeSlots = Convert-LogonHoursToTimeSlots -logonHours $firstUserLogonHours
            $isAllowAll = $firstUserLogonHours -and $firstUserLogonHours.Where({ $_ -eq 0xFF }).Count -eq 21

            # 更新UI状态
            $chkAllowAllTime.Checked = $isAllowAll
            $lstLogonTime.Enabled = -not $isAllowAll
            $script:weekDayCheckboxes | ForEach-Object { $_.Enabled = -not $isAllowAll }

            # 勾选现有允许的时间段
            $lstLogonTime.ClearSelected()
            for ($i = 0; $i -lt $lstLogonTime.Items.Count; $i++) {
                $item = $lstLogonTime.Items[$i].ToString()
                if ($allowedTimeSlots -contains $item) {
                    $lstLogonTime.SetItemChecked($i, $true)
                }
            }

            # 同步更新工作日全选框状态
            foreach ($day in $script:weekDaysUI) {
                $dayItems = $lstLogonTime.Items | Where-Object { $_ -like "$day(*" }
                $checkedDayItems = $dayItems | Where-Object { $lstLogonTime.GetItemChecked($lstLogonTime.Items.IndexOf($_)) }
                $chkDay = $script:weekDayCheckboxes | Where-Object { $_.Text -eq $day }
                $chkDay.Checked = ($checkedDayItems.Count -eq $dayItems.Count) -and ($dayItems.Count -gt 0)
            }

            # 提示多用户设置不一致
            if ($hasDifferentSettings) {
                [System.Windows.Forms.MessageBox]::Show(
                    "选中的用户现有登录时间设置不一致，将统一按本次选择重新设置", 
                    "提示", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("加载登录时间失败：$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $restrictTimeForm.Close()
        }
    }

    # 12. 窗口加载时执行
    $restrictTimeForm.Add_Shown({ 
        LoadUserLogonTime 
    })

    # 13. 保存登录时间限制
    $btnSaveTime.Add_Click({
        $isAllowAll = $chkAllowAllTime.Checked
        $selectedTimeSlots = @()
        if (-not $isAllowAll) {
            # 收集UI选中的时间段（新格式）
            for ($i = 0; $i -lt $lstLogonTime.Items.Count; $i++) {
                if ($lstLogonTime.GetItemChecked($i)) {
                    $selectedTimeSlots += $lstLogonTime.Items[$i].ToString()
                }
            }
        }

        # 生成确认消息
        if ($isAllowAll) {
            $confirmMsg = "将为 $($selectedUsers.Count) 个用户设置【允许所有时间登录】`n`n确定保存吗？"
        } elseif ($selectedTimeSlots.Count -eq 0) {
            $confirmMsg = "将为 $($selectedUsers.Count) 个用户设置【禁止所有时间登录】`n`n确定保存吗？"
        } else {
            $confirmMsg = "将为 $($selectedUsers.Count) 个用户设置允许登录的时间：`n"
            foreach ($day in $script:weekDaysUI) {
                $daySlots = $selectedTimeSlots | Where-Object { $_ -like "$day(*" }
                if ($daySlots) {
                    $confirmMsg += "$day：$($daySlots -join "、")`n"
                }
            }
            $confirmMsg += "`n确定保存吗？"
        }

        # 确认保存
        if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, "确认保存", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne "Yes") {
            return
        }

        try {
            $successUsers = @()
            $failUsers = @()

            # 生成AD需要的logonHours字节数组
            $logonHours = if ($isAllowAll) {
                # 允许所有时间：21字节全为0xFF
                $bytes = New-Object byte[] 21
                for ($i = 0; $i -lt 21; $i++) { $bytes[$i] = 0xFF }
                $bytes
            } else {
                Convert-TimeSlotsToLogonHours -selectedTimeSlots $selectedTimeSlots
            }

            # 批量应用到选中用户
            foreach ($row in $selectedUsers) {
                $displayName = $row.Cells["DisplayName"].Value
                $samAccountName = $row.Cells["SamAccountName"].Value
                $displayName = if ([string]::IsNullOrEmpty($displayName)) { "未设置显示名" } else { $displayName }
                $userLabel = "$displayName（$samAccountName）"

                try {
                    $bytesToSet = [byte[]]$logonHours

                    # 远程执行保存，输出远程执行结果
                    $remoteResult = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                        param($sam, $hours)
                        Import-Module ActiveDirectory -ErrorAction Stop
                        $hoursBytes = [byte[]]$hours
                        
                        Set-ADUser -Identity $sam -Replace @{logonHours = $hoursBytes} -ErrorAction Stop
                        
                        # 验证：设置后立即查询，确认是否生效
                        $updatedUser = Get-ADUser -Identity $sam -Properties logonHours -ErrorAction Stop
                        $isUpdated = $true
                        if ($updatedUser.logonHours.Length -ne $hoursBytes.Length) {
                            $isUpdated = $false
                        } else {
                            for ($i = 0; $i -lt $hoursBytes.Length; $i++) {
                                if ($updatedUser.logonHours[$i] -ne $hoursBytes[$i]) {
                                    $isUpdated = $false; break
                                }
                            }
                        }
                        return $isUpdated
                    } -ArgumentList $samAccountName, $bytesToSet -ErrorAction Stop

                    if ($remoteResult) {
                        $successUsers += $userLabel
                    } else {
                        $failUsers += "$userLabel - 远程设置后验证不一致"
                    }
                } catch {
                    $failMsg = $_.Exception.Message
                    $failUsers += "$userLabel - $failMsg"
                }
            }

            # 显示结果
            $resultTitle = if ($failUsers.Count -eq 0) { "成功" } else { "结果" }
            $resultIcon = if ($failUsers.Count -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }

            $resultMsg = "登录时间限制设置完成！`n"
            $resultMsg += "`n成功用户（共 $($successUsers.Count) 个）："
            $resultMsg += if ($successUsers.Count -eq 0) { "`n无" } else { "`n$($successUsers -join "`n")" }
            $resultMsg += "`n`n失败用户（共 $($failUsers.Count) 个）："
            $resultMsg += if ($failUsers.Count -eq 0) { "`n无" } else { "`n$($failUsers -join "`n")" }

            [System.Windows.Forms.MessageBox]::Show($resultMsg, $resultTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

            # 全部成功则关闭窗口
            if ($failUsers.Count -eq 0) {
                $restrictTimeForm.Close()
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("保存异常：$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # 14. 取消按钮逻辑
    $btnCancelTime.Add_Click({
        $restrictTimeForm.Close()
    })

    # 15. 加载窗口
    $restrictTimeForm.Controls.Add($mainTable)
    $restrictTimeForm.ShowDialog() | Out-Null
}
<# 
OU操作核心函数
#>

function LoadOUList {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    try {
        $script:connectionStatus = "正在从域控读取OU列表..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 远程获取所有OU（包含域根下的OU）
        $script:allOUs = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            Import-Module ActiveDirectory -ErrorAction Stop
            Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName |
                Where-Object { $_.Name -ne "Domain Controllers" } |			
                Select-Object Name, DistinguishedName |
                Sort-Object Name
        } -ErrorAction Stop

        # 处理OU层次结构，生成带层次的显示名称（支持域根下的OU）
        $script:allOUs = $script:allOUs | ForEach-Object {
            $dn = $_.DistinguishedName
            $ouParts = @()
            # 提取DN中的所有OU组件（忽略DC部分）
            $dn -split ',' | ForEach-Object {
                if ($_ -match '^OU=(.+)') {
                    $ouParts += $matches[1]
                }
            }
            # 反转OU组件顺序（DN中是从子到父，反转后为从父到子）
            $hierarchyParts = $ouParts[($ouParts.Count - 1)..0]
            $displayHierarchy = if ($hierarchyParts.Count -gt 1) {
                $hierarchyParts -join ' > '  # 多层级用箭头连接
            }
            else {
                $_.Name  # 顶级OU（域根下的OU）直接显示名称
            }
            [PSCustomObject]@{
                Name              = $_.Name
                DistinguishedName = $_.DistinguishedName
                DisplayHierarchy  = $displayHierarchy  # 层次化显示名称
            }
        } | Sort-Object DisplayHierarchy  # 按层次结构排序

        return $script:allOUs
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "读取OU列表失败：$errorMsg"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("读取OU列表失败：$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
}

# 切换OU组织（支持域根和Users容器）
function SwitchOU {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 加载OU列表（含层次结构）
    $ous = LoadOUList
    if (-not $ous -or $ous.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("未找到任何OU组织", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 获取默认Users容器和域信息（修复域根解析逻辑）
    $defaultUsersOU = $null
    $domainDN = $null
    
    # 优先从domainContext获取域根
    if ($script:domainContext -and $script:domainContext.DomainInfo) {
        $domainDN = $script:domainContext.DomainInfo.DefaultPartition
    }
    # 从当前OU提取域根（关键修复：支持多层子OU）
    if (-not $domainDN -and $script:currentOU) {
        # 正则匹配DN中最后的所有DC组件（域根），例如从"OU=子,OU=父,DC=domain,DC=com"中提取"DC=domain,DC=com"
        if ($script:currentOU -match '(DC=.+)$') {
            $domainDN = $matches[1]
        }
    }
    
    # 生成正确的Users容器路径
    if ($domainDN) {
        $defaultUsersOU = "CN=Users,$domainDN"
        $script:allUsersOU = $defaultUsersOU  # 统一Users容器路径，避免重复定义
    }

    # 创建固定选项（包含域根、默认Users，已移除所有Users）
    $fixedItems = @()
    # 添加域根选项
    if ($domainDN) {
        $fixedItems += [PSCustomObject]@{
            Name              = "域根"
            DistinguishedName = $domainDN
            DisplayHierarchy  = "域根 ($($domainDN -replace 'DC=','.' -replace ',',''))"
        }
    }
    # 添加Users容器选项（显示为"默认(CN=Users,DC=bocmodc3,DC=com)"格式）
    if (-not [string]::IsNullOrWhiteSpace($defaultUsersOU)) {
        $fixedItems += [PSCustomObject]@{
            Name              = "默认Users"
            DistinguishedName = $defaultUsersOU
            DisplayHierarchy  = "默认($defaultUsersOU)"  # 关键修改：显示完整DN
        }
    }

    # 合并固定选项和层次化OU列表
    $displayItems = $fixedItems + $ous

    # 创建OU选择对话框
    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "选择OU组织"
    $ouForm.Size = New-Object System.Drawing.Size(500, 350)  # 加宽窗口以显示完整路径
    $ouForm.StartPosition = "CenterScreen"
    $ouForm.MaximizeBox = $false
    $ouForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

    # 按钮面板
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = "Bottom"
    $buttonPanel.Height = 40
    $buttonPanel.Padding = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
    $buttonPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

    # 列表框（显示层次结构）
    $ouListBox = New-Object System.Windows.Forms.ListBox
    $ouListBox.Dock = "Fill"
    $ouListBox.DisplayMember = "DisplayHierarchy"  # 显示层次化名称
    $ouListBox.ValueMember = "DistinguishedName"
    $ouListBox.Items.AddRange($displayItems)
    $ouListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ouListBox.Font = New-Object System.Drawing.Font("微软雅黑", 9)  # 清晰字体
    if ($script:currentOU) {
        $selectedItem = $ouListBox.Items | Where-Object { $_.DistinguishedName -eq $script:currentOU }
        if ($selectedItem) {
            $ouListBox.SelectedItem = $selectedItem
        }
    }

    # 确定按钮
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "确定"
    $okButton.Width = 100
    $okButton.Height = 30
    $okButton.FlatAppearance.BorderSize = 1	
    $okButton.Location = New-Object System.Drawing.Point(130, 5)
    $okButton.Add_Click({
        if ($ouListBox.SelectedItem) {
            $selectedItem = $ouListBox.SelectedItem
            $script:currentOU = $selectedItem.DistinguishedName
            $script:textOU.Text = $script:currentOU
            
            # 移除所有Users相关逻辑（因选项已删除）
            $script:allUsersOU = $null

            # 清空搜索框
            $script:textSearch.Text = ""
            $script:textGroupSearch.Text = ""
            
            $ouForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        }
    })

    # 取消按钮
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "取消"
    $cancelButton.Width = 100
    $cancelButton.Height = 30
    $cancelButton.FlatAppearance.BorderSize = 1
    $cancelButton.Location = New-Object System.Drawing.Point(255, 5)
    $cancelButton.Add_Click({
        $ouForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    })

    # 添加控件
    $buttonPanel.Controls.Add($okButton)
    $buttonPanel.Controls.Add($cancelButton)
    $ouForm.Controls.Add($ouListBox)
    $ouForm.Controls.Add($buttonPanel)

    if ($ouForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        # 提取显示名称（带层次）
        $displayName = if ($selectedItem.DisplayHierarchy) {
            $selectedItem.DisplayHierarchy
        } else {
            $script:currentOU.Split(',')[0] -replace 'CN=', ''
        }
        $script:connectionStatus = "已切换到OU：$displayName"
        UpdateStatusBar
        
        # 刷新列表
        LoadUserList
        LoadGroupList
    }
}

# 新建OU组织
function CreateNewOU {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("未选择当前OU，请先切换到目标OU后再操作", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $newOUName = [Microsoft.VisualBasic.Interaction]::InputBox(
		"请输入新OU的名称（示例：ITDepartment、财务部）`n`n注意：不可包含特殊字符（/\=+:*#$@?!~`"<>|）", 
		"新建OU组织", 
		""
    )

    # 基础校验
	if ([string]::IsNullOrEmpty($newOUName)) {
		return
	}
	elseif ([string]::IsNullOrWhiteSpace($newOUName)) {
		[System.Windows.Forms.MessageBox]::Show("OU名称不能为空或仅包含空格！", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
		return
	}
		
    # 特殊字符校验
    $invalidChars = '[\\/=+:*#$@?!~"<>|]'
    if ($newOUName -match $invalidChars) {
        $matchedChar = $matches[0]
        [System.Windows.Forms.MessageBox]::Show("OU名称包含非法字符：`"$matchedChar`"`n请删除后重试！", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 确定父容器
    $parentDN = $null
    $parentDisplay = $null

    # 特殊处理：当前OU是CN=Users,DC=bocmodc3,DC=com时，父容器为根域DC=bocmodc3,DC=com
    if ($script:currentOU -eq "CN=Users,DC=bocmodc3,DC=com") {
        $parentDN = "DC=bocmodc3,DC=com"
        $parentDisplay = "域根 (bocmodc3.com)"
    }
    # 其他情况：直接以当前OU作为父容器
    else {
        $parentDN = $script:currentOU
        
        # 生成父容器友好显示名称
        $ouParts = @()
        $parentDN -split ',' | ForEach-Object {
            if ($_ -match '^OU=(.+)') { $ouParts += $matches[1] }
            elseif ($_ -match '^CN=Users') { $ouParts += "Users容器" }
            elseif ($_ -match '^DC=.+') { $ouParts += "域根" }
        }
        $parentDisplay = if ($ouParts.Count -gt 0) {
            $ouParts[($ouParts.Count - 1)..0] -join ' > '
        } else {
            $parentDN.Split(',')[0] -replace '^(OU|CN)=', ''
        }
    }

    # 构建新OU完整路径
    $newOUFullDN = "OU=$newOUName,$parentDN"
    
    # 确认创建
    $confirmMsg = "确认创建OU？`n父容器: $parentDisplay`n新OU名称: $newOUName`n完整路径: $newOUFullDN`n`n注意：此操作将在选定的父容器下创建OU"
    $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "确认新建", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $script:connectionStatus = "正在创建OU：$newOUName..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 远程创建OU
        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($ouName, $parentDN)
            Import-Module ActiveDirectory -ErrorAction Stop
            New-ADOrganizationalUnit -Name $ouName -Path $parentDN -ProtectedFromAccidentalDeletion $true -ErrorAction Stop
        } -ArgumentList $newOUName, $parentDN -ErrorAction Stop

        # 短暂延迟确保AD操作完成同步
        Start-Sleep -Milliseconds 500

        # 创建成功处理
        $script:currentOU = $newOUFullDN
        $script:textOU.Text = $script:currentOU

        $script:connectionStatus = "OU创建成功：$newOUFullDN"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU创建成功！`n完整路径：$newOUFullDN", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # 强制刷新OU列表
        LoadOUList | Out-Null

        # 在UI线程上刷新用户和组列表
        if ($script:mainForm.InvokeRequired) {
            # 明确指定委托类型为Action
            $script:mainForm.Invoke([System.Action]{
                try {
                    LoadUserList
                    LoadGroupList
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("刷新列表失败: $($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            })
        }
        else {
            try {
                LoadUserList
                LoadGroupList
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("刷新列表失败: $($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }

    } catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "already exists") {
            $errorMsg = "OU已存在！请更换名称（如：$newOUName_2）"
        } elseif ($errorMsg -match "permission") {
            $errorMsg = "权限不足！请确认管理员账号拥有在该容器下创建OU的权限"
        } elseif ($errorMsg -match "Path") {
            $errorMsg = "父容器路径错误：$parentDN，请检查该容器是否存在"
        } elseif ($errorMsg -match "invalid DN syntax") {
            $errorMsg = "路径语法错误：$newOUFullDN，请检查名称是否包含特殊字符"
        }
        $script:connectionStatus = "OU创建失败：$errorMsg"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


# 重命名OU组织
function RenameExistingOU {
    [CmdletBinding()]
    param()

    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("未获取到当前OU信息，请重新连接域控！", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 定义不允许重命名的系统关键容器
    $protectedContainers = @(
        "CN=Users,DC=bocmodc3,DC=com",  # 默认Users容器
        "DC=bocmodc3,DC=com"            # 域根容器
    )

    # 检查当前OU是否为系统关键容器
    if ($protectedContainers -contains $script:currentOU) {
        $containerName = if ($script:currentOU -eq "DC=bocmodc3,DC=com") {
            "域根容器（$($script:currentOU -replace 'DC=','.' -replace ',','')）"
        } else {
            "默认Users容器（CN=Users）"
        }
        [System.Windows.Forms.MessageBox]::Show(
            "$containerName 是系统关键容器，不允许重命名！", 
            "操作禁止", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Stop
        )
        return
    }

    # 解析当前OU的名称和父路径
    if ($script:currentOU -match '^OU=(.+?),(.+)') {
        $currentOUName = $matches[1]
        $parentDN = $matches[2]
    } else {
        [System.Windows.Forms.MessageBox]::Show("当前选中的不是有效的OU对象！", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 显示当前OU信息
    $ouParts = @()
    $script:currentOU -split ',' | ForEach-Object {
        if ($_ -match '^OU=(.+)') { $ouParts += $matches[1] }
    }
    $displayHierarchy = if ($ouParts.Count -gt 0) {
        $ouParts[($ouParts.Count - 1)..0] -join ' > '
    } else {
        $currentOUName
    }

    # 获取新名称
    $newOUName = [Microsoft.VisualBasic.Interaction]::InputBox(
        "请输入新的OU名称`n`n当前OU：$displayHierarchy`n`n当前名称：$currentOUName`n`n注意：不可包含特殊字符（/\=+:*#$@?!~`"<>|）", 
        "重命名OU组织", 
        ""
    )

    # 基础校验
    if ([string]::IsNullOrEmpty($newOUName)) {
        return
    }
    elseif ([string]::IsNullOrWhiteSpace($newOUName)) {
        [System.Windows.Forms.MessageBox]::Show("OU名称不能为空或仅包含空格！", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    # 名称未变更
    elseif ($newOUName -eq $currentOUName) {
        [System.Windows.Forms.MessageBox]::Show("新名称与当前名称相同，无需修改！", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 特殊字符校验
    $invalidChars = '[\\/=+:*#$@?!~"<>|]'
    if ($newOUName -match $invalidChars) {
        $matchedChar = $matches[0]
        [System.Windows.Forms.MessageBox]::Show("OU名称包含非法字符：`"$matchedChar`"`n请删除后重试！", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 确认重命名
    $newOUFullDN = "OU=$newOUName,$parentDN"
    $confirmMsg = "确认重命名OU？`n当前OU：$displayHierarchy`n原名称：$currentOUName`n新名称：$newOUName`n新完整路径：$newOUFullDN"
    $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "确认重命名", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $script:connectionStatus = "正在重命名OU：$currentOUName -> $newOUName..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 远程执行重命名
        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($oldDN, $newName)
            Import-Module ActiveDirectory -ErrorAction Stop
            # 使用Rename-ADObject重命名OU
            Rename-ADObject -Identity $oldDN -NewName $newName -ErrorAction Stop
        } -ArgumentList $script:currentOU, $newOUName -ErrorAction Stop

        # 重命名成功处理
        $script:currentOU = $newOUFullDN  # 更新当前OU路径
        $script:textOU.Text = $script:currentOU
        
        $script:connectionStatus = "OU重命名成功：$currentOUName -> $newOUName"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU重命名成功！`n新完整路径：$newOUFullDN", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
		Start-Sleep -Milliseconds 500
        # 刷新相关列表
        LoadOUList
        if ($script:mainForm.InvokeRequired) {
            # 明确指定委托类型为Action
            $script:mainForm.Invoke([System.Action]{
                try {
                    LoadUserList
                    LoadGroupList
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("刷新列表失败: $($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            })
        }
        else {
            try {
                LoadUserList
                LoadGroupList
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("刷新列表失败: $($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }

    } catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "already exists") {
            $errorMsg = "名称已存在！该父容器下已有名为`"$newOUName`"的OU"
        } elseif ($errorMsg -match "permission") {
            $errorMsg = "权限不足！请确认管理员账号拥有OU重命名权限"
        } elseif ($errorMsg -match "not found") {
            $errorMsg = "OU不存在！可能已被删除或路径错误"
        } elseif ($errorMsg -match "invalid name") {
            $errorMsg = "无效的名称格式！请检查是否包含不支持的字符或格式"
        }
        
        $script:connectionStatus = "OU重命名失败：$errorMsg"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}



# 删除OU组织
function DeleteExistingOU {
    [CmdletBinding()]
    param()

    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("未获取到当前OU信息，请重新连接域控！", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 定义不允许删除的系统关键容器
    $protectedContainers = @(
        "CN=Users,DC=bocmodc3,DC=com",  # 默认Users容器
        "DC=bocmodc3,DC=com"            # 域根容器
    )

    # 检查当前OU是否为系统关键容器
    if ($protectedContainers -contains $script:currentOU) {
        $containerName = if ($script:currentOU -eq "DC=bocmodc3,DC=com") {
            "域根容器（$($script:currentOU -replace 'DC=','.' -replace ',','')）"
        } else {
            "默认Users容器（CN=Users）"
        }
        [System.Windows.Forms.MessageBox]::Show(
            "$containerName 是系统关键容器，不允许删除！`n这是域的基础组件，删除会导致域功能异常。", 
            "操作禁止", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Stop
        )
        return
    }

    # 显示OU层次信息（支持域根下的OU）
    $ouParts = @()
    $script:currentOU -split ',' | ForEach-Object {
        if ($_ -match '^OU=(.+)') { $ouParts += $matches[1] }
    }
    $displayHierarchy = if ($ouParts.Count -gt 0) {
        $ouParts[($ouParts.Count - 1)..0] -join ' > '
    } else {
        $script:currentOU.Split(',')[0] -replace 'OU=',''  # 域根下的顶级OU
    }
    $deleteMsg = "警告：删除OU会同时删除其下所有对象（用户、组、子OU）！`n当前OU层次：$displayHierarchy`n完整路径：$script:currentOU`n`n确定要删除吗？"

    $confirmResult = [System.Windows.Forms.MessageBox]::Show($deleteMsg, "高危操作确认", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $script:connectionStatus = "正在删除OU：$displayHierarchy..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 远程删除OU（含子对象）
        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($ouDN)
            Import-Module ActiveDirectory -ErrorAction Stop
            
            # 解除保护
            try {
                Set-ADOrganizationalUnit -Identity $ouDN -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
            }
            catch {
                # 针对系统容器的特殊提示
                if ($ouDN -in "CN=Users,DC=bocmodc3,DC=com", "DC=bocmodc3,DC=com") {
                    Write-Error "系统关键容器不允许解除保护"
                } else {
                    Write-Warning "解除保护时出错: $($_.Exception.Message)"
                }
            }
            
            # 递归删除
            Remove-ADOrganizationalUnit -Identity $ouDN -Recursive -Confirm:$false -ErrorAction Stop
        } -ArgumentList $script:currentOU -ErrorAction Stop

        # 短暂延迟确保AD操作完成同步
        Start-Sleep -Milliseconds 500

        # 删除成功处理
        $script:connectionStatus = "OU删除成功：$displayHierarchy"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU已彻底删除（含所有子对象）！", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # 重置当前OU为父容器（支持域根和Users容器）
        $parentDN = $script:currentOU -replace '^[^,]+,', ''
        if ($parentDN -match '^OU=|^CN=Users,|^DC=') {  # 父容器可以是OU、Users容器或域根
            $script:currentOU = $parentDN
        }
        else {
            # 父容器是域根时，切换到域根
            $script:currentOU = $parentDN
        }
        $script:textOU.Text = $script:currentOU

        # 强制刷新OU列表
        LoadOUList | Out-Null

        # 在UI线程上刷新用户和组列表
        if ($script:mainForm.InvokeRequired) {
            $script:mainForm.Invoke([System.Action]{
                try {
                    LoadUserList
                    LoadGroupList
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("刷新列表失败: $($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            })
        }
        else {
            try {
                LoadUserList
                LoadGroupList
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("刷新列表失败: $($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }

    } catch {
        $errorMsg = $_.Exception.Message
        # 针对系统关键容器的错误提示优化
        if ($script:currentOU -eq "CN=Users,DC=bocmodc3,DC=com") {
            $errorMsg = "无法删除默认Users容器！这是系统内置容器，不允许删除。"
        } elseif ($script:currentOU -eq "DC=bocmodc3,DC=com") {
            $errorMsg = "无法删除域根容器！这是域的基础，删除会导致整个域崩溃。"
        } elseif ($errorMsg -match "not found") {
            $errorMsg = "OU不存在！可能已被删除或路径错误"
        } elseif ($errorMsg -match "permission") {
            $errorMsg = "权限不足！请确认管理员账号拥有OU删除权限"
        } elseif ($errorMsg -match "保护") {
            $errorMsg = "无法解除OU保护！可能该OU被系统锁定或需要更高权限"
        }
        
        Write-Error $errorMsg
        $script:connectionStatus = "OU删除失败：$errorMsg"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


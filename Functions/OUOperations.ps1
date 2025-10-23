<# 
OU操作核心函数
#>


function Get-DomainRootDN {
    # 1. 检查基础依赖（必须先连接域控）
    if (-not $script:remoteSession -and -not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控服务器", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return $null
    }

    try {
        # 2. 优先从远程会话获取域信息（最可靠，直接读取域控配置）
        $domainInfo = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            Import-Module ActiveDirectory -ErrorAction Stop
            # 获取当前域的完整信息，DefaultPartition即为域根DN（如DC=contoso,DC=com）
            $adDomain = Get-ADDomain -ErrorAction Stop
            return @{
                DomainRootDN = $adDomain.DefaultPartition  # 核心：域根DN
                DomainDNS    = $adDomain.DNSRoot           # 辅助：域DNS名称（如contoso.com）
            }
        } -ErrorAction Stop

        # 3. 验证并返回域根DN
        if (-not [string]::IsNullOrWhiteSpace($domainInfo.DomainRootDN)) {
            return $domainInfo.DomainRootDN
        }

        # 4. 备用方案：从domainContext提取（兼容旧逻辑）
        if ($script:domainContext -and $script:domainContext.DomainInfo -and $script:domainContext.DomainInfo.DefaultPartition) {
            return $script:domainContext.DomainInfo.DefaultPartition
        }

        # 5. 兜底方案：从当前OU反向解析（仅当currentOU已存在时）
        if ($script:currentOU -match '(DC=.+)$') {
            return $matches[1]
        }

        # 6. 所有方案失败
        [System.Windows.Forms.MessageBox]::Show("无法获取域根信息，请重新连接域控", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
    catch {
        $errorMsg = "获取域根失败：$($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
}

# 加载OU列表
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

#切换OU组织
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

    # 关键修改：动态获取域根和默认Users容器
    $domainRootDN = Get-DomainRootDN  # 调用通用函数获取域根
    if (-not $domainRootDN) { return }  # 获取失败则终止

    $defaultUsersOU = "CN=Users,$domainRootDN"  # 动态生成Users容器路径
    $script:allUsersOU = $defaultUsersOU  # 统一Users容器路径

    # 创建固定选项（域根、默认Users）
    $fixedItems = @()
    # 添加域根选项（显示格式：域根 (contoso.com)）
    $domainDNS = $domainRootDN -replace 'DC=','.' -replace ',',''  # 把DC=domain,DC=com转成domain.com
    $fixedItems += [PSCustomObject]@{
        Name              = "域根"
        DistinguishedName = $domainRootDN
        DisplayHierarchy  = "域根 ($domainDNS)"
    }
    # 添加Users容器选项（显示完整动态路径）
    $fixedItems += [PSCustomObject]@{
        Name              = "默认Users"
        DistinguishedName = $defaultUsersOU
        DisplayHierarchy  = "默认($defaultUsersOU)"
    }

    # 合并固定选项和层次化OU列表
    $displayItems = $fixedItems + $ous

    # 创建OU选择对话框（逻辑不变，仅数据源改为动态）
    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "选择OU组织"
    $ouForm.Size = New-Object System.Drawing.Size(500, 350)
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
    $ouListBox.DisplayMember = "DisplayHierarchy"
    $ouListBox.ValueMember = "DistinguishedName"
    $ouListBox.Items.AddRange($displayItems)
    $ouListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ouListBox.Font = New-Object System.Drawing.Font("微软雅黑", 9)
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
            
            $script:allUsersOU = $null  # 清空冗余变量

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

#新建OU组织
function CreateNewOU {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("未选择当前OU，请先切换到目标OU后再操作", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    #获取域根和默认Users容器
    $domainRootDN = Get-DomainRootDN
    if (-not $domainRootDN) { return }
    $defaultUsersOU = "CN=Users,$domainRootDN"

    # 输入新OU名称（逻辑不变）
    $newOUName = [Microsoft.VisualBasic.Interaction]::InputBox(
		"请输入新OU的名称（示例：ITDepartment、财务部）`n`n注意：不可包含特殊字符（/\=+:*#$@?!~`"<>|）", 
		"新建OU组织", 
		""
    )

    # 基础校验（逻辑不变）
	if ([string]::IsNullOrEmpty($newOUName)) {
		return
	}
	elseif ([string]::IsNullOrWhiteSpace($newOUName)) {
		[System.Windows.Forms.MessageBox]::Show("OU名称不能为空或仅包含空格！", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
		return
	}
		
    # 特殊字符校验（逻辑不变）
    $invalidChars = '[\\/=+:*#$@?!~"<>|]'
    if ($newOUName -match $invalidChars) {
        $matchedChar = $matches[0]
        [System.Windows.Forms.MessageBox]::Show("OU名称包含非法字符：`"$matchedChar`"`n请删除后重试！", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 判断父容器
    $parentDN = $null
    $parentDisplay = $null

    # 当前OU是【默认Users容器】时，父容器为域根
    if ($script:currentOU -eq $defaultUsersOU) {
        $parentDN = $domainRootDN
        $domainDNS = $domainRootDN -replace 'DC=','.' -replace ',',''
        $parentDisplay = "域根 ($domainDNS)"
    }
    # 其他情况：直接以当前OU作为父容器
    else {
        $parentDN = $script:currentOU
        
        # 生成父容器友好显示名称（逻辑不变）
        $ouParts = @()
        $parentDN -split ',' | ForEach-Object {
            if ($_ -match '^OU=(.+)') { $ouParts += $matches[1] }
            elseif ($_ -eq "CN=Users") { $ouParts += "Users容器" }
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

        # 短暂延迟确保AD同步
        Start-Sleep -Milliseconds 500

        # 创建成功处理
        $script:currentOU = $newOUFullDN
        $script:textOU.Text = $script:currentOU

        $script:connectionStatus = "OU创建成功：$newOUFullDN"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU创建成功！`n完整路径：$newOUFullDN", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # 强制刷新OU列表
        LoadOUList | Out-Null

        # 刷新用户和组列表
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
        # 错误处理
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

#重命名OU组织
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

    # 获取受保护容器（域根、默认Users）
    $domainRootDN = Get-DomainRootDN
    if (-not $domainRootDN) { return }
    $defaultUsersOU = "CN=Users,$domainRootDN"
    $protectedContainers = @(
        $defaultUsersOU,  # 动态默认Users容器
        $domainRootDN     # 动态域根容器
    )

    # 检查当前OU是否为系统关键容器
    if ($protectedContainers -contains $script:currentOU) {
        $containerName = if ($script:currentOU -eq $domainRootDN) {
            $domainDNS = $domainRootDN -replace 'DC=','.' -replace ',',''
            "域根容器（$domainDNS）"
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
            Rename-ADObject -Identity $oldDN -NewName $newName -ErrorAction Stop
        } -ArgumentList $script:currentOU, $newOUName -ErrorAction Stop

        # 重命名成功处理
        $script:currentOU = $newOUFullDN
        $script:textOU.Text = $script:currentOU
        
        $script:connectionStatus = "OU重命名成功：$currentOUName -> $newOUName"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU重命名成功！`n新完整路径：$newOUFullDN", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
		Start-Sleep -Milliseconds 500
        # 刷新相关列表
        LoadOUList
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
        # 错误处理
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

#删除OU组织
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

    # 获取受保护容器（域根、默认Users）
    $domainRootDN = Get-DomainRootDN
    if (-not $domainRootDN) { return }
    $defaultUsersOU = "CN=Users,$domainRootDN"
    $protectedContainers = @(
        $defaultUsersOU,  # 动态默认Users容器
        $domainRootDN     # 动态域根容器
    )

    # 检查当前OU是否为系统关键容器
    if ($protectedContainers -contains $script:currentOU) {
        $containerName = if ($script:currentOU -eq $domainRootDN) {
            $domainDNS = $domainRootDN -replace 'DC=','.' -replace ',',''
            "域根容器（$domainDNS）"
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

    # 显示OU层次信息
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
                Write-Warning "解除保护时出错: $($_.Exception.Message)"
            }
            
            # 递归删除
            Remove-ADOrganizationalUnit -Identity $ouDN -Recursive -Confirm:$false -ErrorAction Stop
        } -ArgumentList $script:currentOU -ErrorAction Stop

        # 短暂延迟确保AD同步
        Start-Sleep -Milliseconds 500

        # 删除成功处理
        $script:connectionStatus = "OU删除成功：$displayHierarchy"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU已彻底删除（含所有子对象）！", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # 重置当前OU为父容器
        $parentDN = $script:currentOU -replace '^[^,]+,', ''
        if ($parentDN -match '^OU=|^CN=Users,|^DC=') {  # 父容器可以是OU、Users容器或域根
            $script:currentOU = $parentDN
        }
        else {
            $script:currentOU = $parentDN
        }
        $script:textOU.Text = $script:currentOU

        # 强制刷新OU列表
        LoadOUList | Out-Null

        # 刷新用户和组列表
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
        # 适配错误提示
        $errorMsg = $_.Exception.Message
        if ($script:currentOU -eq $defaultUsersOU) {
            $errorMsg = "无法删除默认Users容器！这是系统内置容器，不允许删除。"
        } elseif ($script:currentOU -eq $domainRootDN) {
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
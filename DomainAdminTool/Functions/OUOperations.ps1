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

        # 远程获取所有OU
        $script:allOUs = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            Import-Module ActiveDirectory -ErrorAction Stop
            Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName |
                Where-Object { $_.Name -ne "Domain Controllers" } |			
                Select-Object Name, DistinguishedName |
                Sort-Object Name
        } -ErrorAction Stop

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



# 切换OU组织
function SwitchOU {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 加载OU列表
    $ous = LoadOUList
    if (-not $ous -or $ous.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("未找到任何OU组织", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 获取默认Users容器和所有Users容器信息
    $defaultUsersOU = $null
    $allUsersOU = $null
    $domainDN = $null
    
    # 尝试获取域信息
    if ($script:domainContext -and $script:domainContext.DomainInfo) {
        $domainDN = $script:domainContext.DomainInfo.DefaultPartition
    }
    # 从当前OU解析域信息
    if (-not $domainDN -and $script:currentOU) {
        $domainDN = $script:currentOU -replace '^[^,]+,', ''
    }
    
    # 构建默认Users容器DN和所有Users容器DN
    if ($domainDN) {
        $defaultUsersOU = "CN=Users,$domainDN"
        $script:allUsersOU = "CN=Users,$domainDN"
    }

    # 创建固定选项对象并添加到列表
    $fixedItems = @()
    if (-not [string]::IsNullOrWhiteSpace($defaultUsersOU)) {
        $fixedItems += [PSCustomObject]@{
            Name = "默认Users"
            DistinguishedName = $defaultUsersOU
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($script:allUsersOU)) {
        $fixedItems += [PSCustomObject]@{
            Name = "所有Users"
            DistinguishedName = $script:allUsersOU
        }
    }

    # 合并固定选项和原有OU列表（固定选项放在前面）
    $displayItems = $fixedItems + $ous

    # 创建OU选择对话框
    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "选择OU组织"
    $ouForm.Size = New-Object System.Drawing.Size(350, 250)  # 固定窗口大小
    $ouForm.StartPosition = "CenterScreen"
    $ouForm.MaximizeBox = $false
    $ouForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

    # 创建面板放置按钮（横向排列）
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = "Bottom"
    $buttonPanel.Height = 40
    $buttonPanel.Padding = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
    $buttonPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

    $ouListBox = New-Object System.Windows.Forms.ListBox
    $ouListBox.Dock = "Fill"
    $ouListBox.DisplayMember = "Name"
    $ouListBox.ValueMember = "DistinguishedName"
    $ouListBox.Items.AddRange($displayItems)
    $ouListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
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
    $okButton.Location = New-Object System.Drawing.Point(55, 5)
    $okButton.Add_Click({
        if ($ouListBox.SelectedItem) {
			$selectedItem = $ouListBox.SelectedItem
			$script:currentOU = $selectedItem.DistinguishedName
			$script:textOU.Text = $script:currentOU
			
			# 根据选择设置allUsersOU状态
			if ($selectedItem.Name -eq "所有Users") {
				# 选中"所有Users"时，设置allUsersOU为对应DN
				$script:allUsersOU = $selectedItem.DistinguishedName
			} else {
				# 其他选项时，清空allUsersOU（表示只加载指定OU）
				$script:allUsersOU = $null
			}
			
			$ouForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			}
    })

    # 取消按钮
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "取消"
    $cancelButton.Width = 100
    $cancelButton.Height = 30
	$cancelButton.FlatAppearance.BorderSize = 1
    $cancelButton.Location = New-Object System.Drawing.Point(180, 5)
    $cancelButton.Add_Click({
        $ouForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    })

    # 添加控件到面板和表单
    $buttonPanel.Controls.Add($okButton)
    $buttonPanel.Controls.Add($cancelButton)
    $ouForm.Controls.Add($ouListBox)
    $ouForm.Controls.Add($buttonPanel)

    if ($ouForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $displayName = switch -Wildcard ($script:currentOU) {
            "^CN=Users," { "默认Users容器" }
            "^CN=Users," { "所有Users容器" }
            default { $script:currentOU.Split(',')[0] }
        }
        $script:connectionStatus = "已切换到OU：$displayName"
        UpdateStatusBar
        
        # 切换OU后刷新用户和组列表
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


    $newOUName = [Microsoft.VisualBasic.Interaction]::InputBox(
		"请输入新OU的名称（示例：ITDepartment、Finance）`n`n注意：不可包含特殊字符（/ \ : * ? `" < > |）", 
		"新建OU组织", 
		""
    )


    # 基础空值校验
	if ([string]::IsNullOrEmpty($newOUName)) {
		return
	}
	elseif ([string]::IsNullOrWhiteSpace($newOUName)) {
		[System.Windows.Forms.MessageBox]::Show("OU名称不能为空或仅包含空格！", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
		return
	}


    # 特殊字符正则校验
    $invalidChars = '[\\/:*?"<>|]'
    if ($newOUName -match $invalidChars) {
        $matchedChar = $matches[0]
        [System.Windows.Forms.MessageBox]::Show("OU名称包含非法字符：`"$matchedChar`"`n请删除后重试！", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 自动获取域的默认分区
    $domainDN = $null
    $logMessages = @("开始获取域信息...")
    
    # 方法1: 从domainContext获取（详细诊断）
    if (-not $domainDN) {
        try {
            $logMessages += "尝试从domainContext获取域信息..."
            if ($script:domainContext) {
                $logMessages += "domainContext存在"
                if ($script:domainContext.DomainInfo) {
                    $logMessages += "DomainInfo属性存在"
                    $domainDN = $script:domainContext.DomainInfo.DefaultPartition
                    $logMessages += "从domainContext获取到: $domainDN"
                }
                else {
                    $logMessages += "domainContext中不存在DomainInfo属性"
                }
            }
            else {
                $logMessages += "script:domainContext为空"
            }
        }
        catch {
            $logMessages += "从domainContext获取失败: $($_.Exception.Message)"
        }
    }

    # 方法2: 从远程会话获取（直接构造预期格式）
    if (-not $domainDN -and $script:remoteSession) {
        try {
            $logMessages += "尝试从远程会话获取域信息..."
            # 方法2.1: 使用Get-ADDomain
            $domainInfo = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                Import-Module ActiveDirectory -ErrorAction Stop
                Get-ADDomain -ErrorAction Stop
            } -ErrorAction Stop
            
            if ($domainInfo -and $domainInfo.DefaultPartition) {
                $domainDN = $domainInfo.DefaultPartition
                $logMessages += "从Get-ADDomain获取到: $domainDN"
            }
            else {
                $logMessages += "Get-ADDomain未返回有效信息"
                # 方法2.2: 尝试获取当前域控制器的域名并转换为DN
                $domainController = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                    $env:USERDNSDOMAIN
                } -ErrorAction Stop
                
                if ($domainController) {
                    $logMessages += "获取到域控制器: $domainController"
                    $domainParts = $domainController -split '\.'
                    if ($domainParts.Count -ge 2) {
                        $domainDN = "dc=$($domainParts -join ',dc=')"
                        $logMessages += "转换后得到域DN: $domainDN"
                    }
                }
            }
        }
        catch {
            $logMessages += "从远程会话获取失败: $($_.Exception.Message)"
        }
    }

    # 方法3: 从当前OU解析（如果有当前OU）
    if (-not $domainDN -and -not [string]::IsNullOrWhiteSpace($script:currentOU)) {
        try {
            $logMessages += "尝试从当前OU解析域信息: $script:currentOU"
            # 从当前OU中提取域信息
            $domainDN = $script:currentOU -replace '^[^,]+,', ''
            # 验证是否是有效的域格式
            if ($domainDN -match '^DC=') {
                $logMessages += "从当前OU解析得到: $domainDN"
            }
            else {
                $logMessages += "解析结果格式无效: $domainDN"
                $domainDN = $null
            }
        }
        catch {
            $logMessages += "从当前OU解析失败: $($_.Exception.Message)"
        }
    }

    # 方法4: 尝试从计算机加入的域获取
    if (-not $domainDN) {
        try {
            $logMessages += "尝试从本地计算机加入的域获取..."
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $domainName = $domain.Name
            $logMessages += "当前计算机加入的域: $domainName"
            
            # 转换为DN格式
            $domainParts = $domainName -split '\.'
            if ($domainParts.Count -ge 2) {
                $domainDN = "dc=$($domainParts -join ',dc=')"
                $logMessages += "转换后得到域DN: $domainDN"
            }
        }
        catch {
            $logMessages += "从本地计算机获取域信息失败: $($_.Exception.Message)"
        }
    }

    # 如果所有自动方法都失败，记录日志并提示手动输入
    if (-not $domainDN) {
        # 将日志保存到临时文件以便排查问题
        $logPath = "$env:TEMP\DomainInfoLog.txt"
        $logMessages | Out-File -FilePath $logPath -Encoding utf8
        
        $domainDN = [Microsoft.VisualBasic.Interaction]::InputBox(
            "无法自动获取域信息（日志已保存到: $logPath）`n请手动输入域的DistinguishedName`n例如：DC=abc-test,DC=com", 
            "手动输入域信息", 
            "dc=abc-test,dc=com"  # 预设用户的域
        )
        
        if ([string]::IsNullOrWhiteSpace($domainDN) -or $domainDN -notmatch '^DC=') {
            [System.Windows.Forms.MessageBox]::Show("无效的域信息！格式应为DC=domain,DC=com", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }

    # 最终校验Path参数
    if ([string]::IsNullOrWhiteSpace($domainDN)) {
        [System.Windows.Forms.MessageBox]::Show("无法获取有效的域信息，无法创建OU！", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 确认新OU的父容器
    $newOUFullDN = "OU=$newOUName,$domainDN"
    $confirmMsg = "确认在以下位置创建OU？`n父容器：$domainDN`n新OU完整路径：$newOUFullDN"
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
        } -ArgumentList $newOUName, $domainDN -ErrorAction Stop

        # 创建成功处理
        $script:connectionStatus = "OU创建成功：$newOUFullDN"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU创建成功！`n完整路径：$newOUFullDN", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # 切换到新创建的OU
        $script:currentOU = $newOUFullDN
        $script:textOU.Text = $script:currentOU

        # 刷新用户和组列表
        LoadUserList
        LoadGroupList

    } catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "already exists") {
            $errorMsg = "OU已存在！请更换名称（如：$newOUName_2）"
        } elseif ($errorMsg -match "permission") {
            $errorMsg = "权限不足！请确认管理员账号拥有OU创建权限"
        } elseif ($errorMsg -match "Path") {
            $errorMsg = "路径参数错误：$domainDN，请检查域信息是否正确"
        }
        $script:connectionStatus = "OU创建失败：$errorMsg"
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

    # 校验当前是否有选中的OU
    if ([string]::IsNullOrWhiteSpace($script:currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("未获取到当前OU信息，请重新连接域控！", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 二次确认（高危操作）
    $confirmMsg = "警告：删除OU会同时删除其下所有对象（用户、组、子OU）！`n确定要删除以下OU吗？`n$script:currentOU"
    $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "高危操作确认", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $script:connectionStatus = "正在删除OU：$script:currentOU..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 远程删除OU
        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($ouDN)
            Import-Module ActiveDirectory -ErrorAction Stop
            
            # 解除保护（增加错误处理）
            try {
                Set-ADOrganizationalUnit -Identity $ouDN -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
            }
            catch {
                Write-Warning "解除保护时出错: $($_.Exception.Message)"
            }
            
            # 递归删除OU及所有子对象
            Remove-ADOrganizationalUnit -Identity $ouDN -Recursive -Confirm:$false -ErrorAction Stop
        } -ArgumentList $script:currentOU -ErrorAction Stop

        # 删除成功处理
        $script:connectionStatus = "OU删除成功：$script:currentOU"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU已彻底删除（含子对象）！", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # 重置当前OU为域控默认Users OU（使用更可靠的域信息获取方式）
        $domainDN = $null
        # 尝试从domainContext获取域信息
        if ($script:domainContext -and $script:domainContext.DomainInfo) {
            $domainDN = $script:domainContext.DomainInfo.DefaultPartition
        }
        # 如果失败，尝试从当前OU解析
        if (-not $domainDN -and $script:currentOU) {
            $domainDN = $script:currentOU -replace '^[^,]+,', ''
        }
        
        # 如果获取到域信息，设置默认Users OU
        if ($domainDN) {
            $defaultUsersOU = "CN=Users,$domainDN"
            $script:currentOU = $defaultUsersOU
            $script:textOU.Text = $script:currentOU
        }
        else {
            $script:currentOU = $null
            $script:textOU.Text = ""
        }

        # 刷新用户和组列表
        LoadUserList
        LoadGroupList

    } catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "not found") {
            $errorMsg = "OU不存在！可能已被删除或路径错误"
        } elseif ($errorMsg -match "permission") {
            $errorMsg = "权限不足！请确认管理员账号拥有OU删除权限"
        }
		Write-Error $errorMsg
        $script:connectionStatus = "OU删除失败：$errorMsg"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


    
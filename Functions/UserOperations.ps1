<# 
用户相关核心操作 
#>

function LoadUserList {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    try {
        $script:connectionStatus = "正在加载 OU: $($script:currentOU) 下的用户..."
        UpdateStatusBar
        $script:mainForm.Refresh()
        
        $script:allUsers.Clear()
        $script:filteredUsers.Clear()

        # 远程加载用户（逻辑不变）
        $remoteUsers = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($searchBase, $allUsersOU)
            Import-Module ActiveDirectory -ErrorAction Stop
            if ($allUsersOU) {
                $users = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, MemberOf, EmailAddress, TelephoneNumber, LockedOut, Description, Enabled, AccountExpirationDate -ErrorAction Stop
            } else {
                $users = Get-ADUser -Filter * -SearchBase $searchBase `
                    -Properties DisplayName, SamAccountName, MemberOf, EmailAddress, TelephoneNumber, LockedOut, Description, Enabled, AccountExpirationDate `
                    -ErrorAction Stop
            }

            # 处理用户数据（逻辑不变）
            $users | ForEach-Object {
                $groupNames = $_.MemberOf | ForEach-Object { if ($_ -match 'CN=([^,]+)') { $matches[1] } }
                $groupsString = if ($groupNames) { $groupNames -join ', ' } else { '无' }
                [PSCustomObject]@{
                    DisplayName       = $_.DisplayName
                    SamAccountName    = $_.SamAccountName
                    MemberOf          = $groupsString
                    EmailAddress      = $_.EmailAddress
                    TelePhone         = $_.TelephoneNumber
                    AccountLockout    = [bool]$_.LockedOut
                    Description       = $_.Description
                    Enabled           = [bool]$_.Enabled
                    AccountExpirationDate = $_.AccountExpirationDate
                }
            }			
        } -ArgumentList $script:currentOU, $script:allUsersOU -ErrorAction Stop

        # 填充数据（逻辑不变）
        $remoteUsers | ForEach-Object {
            $null = $script:allUsers.Add($_)
            $null = $script:filteredUsers.Add($_)
        }
        
        # ---------------------- 关键：默认全显配置 ----------------------
        $script:defaultShowAll = $true  # 切换OU后强制默认全显
        $script:currentUserPage = 1  # 重置当前页码为1
        # 全显时总页数=1（用总数据量作为pageSize）
        $script:totalUserPages = Get-TotalPages -totalCount $script:filteredUsers.Count -pageSize $script:filteredUsers.Count  
        
        # 1. 绑定全量数据到DataGridView
        $script:userDataGridView.DataSource = $null
        $script:userDataGridView.DataSource = $script:filteredUsers
        
        # 2. 同步分页控件状态（禁用翻页按钮，显示正确页码）
        $script:lblUserPageInfo.Text = "第 $script:currentUserPage 页 / 共 $script:totalUserPages 页（总计 $($script:filteredUsers.Count) 条）"
        $script:btnUserPrev.Enabled = $false  # 全显时无法上一页
        $script:btnUserNext.Enabled = $false  # 全显时无法下一页
        $script:txtUserJumpPage.Text = "1"  # 跳转框默认显示1
        $script:userPaginationPanel.Visible = $true  # 强制显示分页控件
        # ----------------------------------------------------------------
        
        # 更新状态（逻辑不变）
        $script:userCountStatus = $script:allUsers.Count
        $script:connectionStatus = "已加载 OU: $($script:currentOU) 下的 $($script:userCountStatus) 个用户"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $script:connectionStatus = "加载用户失败: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
    }
}



function ToggleUserEnabled {
    param($rowIndex)
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $script:userDataGridView.CancelEdit()
        return
    }
    
    $selectedRow = $script:userDataGridView.Rows[$rowIndex]
    $enabledCell = $selectedRow.Cells["Enabled"]
    
    # 获取当前状态
    $currentState = $false
    if ($enabledCell.Value -ne $null) { $currentState = [bool]$enabledCell.Value }
    $newState = -not $currentState
    $action = if ($newState) { "启用" } else { "禁用" }
    
    # 获取账号
    $username = $null
    if ($selectedRow.DataBoundItem -ne $null) {
        $username = $selectedRow.DataBoundItem.SamAccountName.ToString().Trim()
    }
    if ([string]::IsNullOrEmpty($username)) {
        $username = $selectedRow.Cells["SamAccountName"].Value.ToString().Trim()
    }
    if ([string]::IsNullOrEmpty($username)) {
        [System.Windows.Forms.MessageBox]::Show("未找到账号信息", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $script:userDataGridView.CancelEdit()
        return
    }
    
    # 确认操作
    if ([System.Windows.Forms.MessageBox]::Show("确定要$($action)账号 [$username] 吗？", "确认", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne 'Yes') {
        $script:userDataGridView.CancelEdit()
        return
    }
    
    try {
        $script:connectionStatus = "正在$($action)账号..."
        UpdateStatusBar
        $script:mainForm.Refresh()
        
        # 远程执行启用/禁用
        $remoteResult = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($username, $newState)
            Import-Module ActiveDirectory -ErrorAction Stop
            $user = Get-ADUser -Filter { SamAccountName -eq $username } -ErrorAction Stop
            Set-ADUser -Identity $user.DistinguishedName -Enabled $newState -ErrorAction Stop
            $updatedUser = Get-ADUser -Identity $user.DistinguishedName -Properties Enabled -ErrorAction Stop
            return $updatedUser.Enabled
        } -ArgumentList $username, $newState -ErrorAction Stop
        
        if ($remoteResult -ne $newState) { throw "操作后状态未变更" }
        [System.Windows.Forms.MessageBox]::Show("账号 [$username] $($action)成功", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "${action}失败: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        $script:userDataGridView.CancelEdit()
        [System.Windows.Forms.MessageBox]::Show("账号 [$username] $($action)失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}



function CreateNewUser {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    # 获取输入并确保基本参数不为空
    $cnName = $script:textCnName.Text.Trim()
    $username = $script:textPinyin.Text.Trim()
    $email = $script:textEmail.Text.Trim()
	$phone = $script:textPhone.Text.Trim()
    $description = $script:textDescription.Text.Trim()
    $password = $script:textNewPassword.Text
    $confirm = $script:textConfirmPassword.Text
    $prefix = $script:textPrefix.Text.Trim()
    $currentOU = $script:currentOU
    $expiryDate = $script:dateExpiry.Value.AddDays(1)
    $domainDNSRoot = $script:domainContext.DomainInfo.DNSRoot
    $neverExpire = $script:chkNeverExpire.Checked

    # 增强校验，确保所有必要参数都有值
    if ([string]::IsNullOrWhiteSpace($cnName)) {
        [System.Windows.Forms.MessageBox]::Show("姓名不能为空", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ([string]::IsNullOrWhiteSpace($username)) {
        [System.Windows.Forms.MessageBox]::Show("账号不能为空", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ([string]::IsNullOrWhiteSpace($password)) {
        [System.Windows.Forms.MessageBox]::Show("密码不能为空", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($password -ne $confirm) {
        [System.Windows.Forms.MessageBox]::Show("两次密码不一致", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($password.Length -lt 8 -or $password -notmatch '[A-Z]' -or $password -notmatch '[a-z]' -or $password -notmatch '[0-9]' -or $password -notmatch '[^a-zA-Z0-9]') {
        [System.Windows.Forms.MessageBox]::Show("密码需≥8位，包含大、小写字母、数字和特殊字符（如@#）", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ([string]::IsNullOrWhiteSpace($currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("请选择用户组织单位(OU)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ([string]::IsNullOrWhiteSpace($domainDNSRoot)) {
        [System.Windows.Forms.MessageBox]::Show("无法获取域信息，请重新连接", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    try {
        $script:connectionStatus = "正在远程创建账号[$username]..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 确保远程会话存在
        if (-not $script:remoteSession -or $script:remoteSession.State -ne 'Opened') {
            throw "远程会话未连接，请重新连接到域控"
        }

        # 远程创建用户
        $remoteParams = @{
            Session = $script:remoteSession
            ScriptBlock = {
                param(
                    [Parameter(Mandatory=$true)]
                    [string]$cnName,
                    [Parameter(Mandatory=$true)]
                    [string]$username,
                    [string]$email,
					[string]$phone,
                    [Parameter(Mandatory=$true)]
                    [string]$NameOU,
                    [string]$description,
                    [Parameter(Mandatory=$true)]
                    [string]$password,
                    [DateTime]$expiryDate,
                    [Parameter(Mandatory=$true)]
                    [string]$domainDNSRoot,
                    [bool]$neverExpire
                )

                # 确保ActiveDirectory模块已加载
                if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
                    throw "ActiveDirectory模块未安装"
                }
                Import-Module ActiveDirectory -ErrorAction Stop

                # 定义复姓列表
				$doubleSurnames = @(
					"欧阳", "太史", "端木", "上官", "司马", "东方", "独孤", "南宫",
					"万俟", "闻人", "夏侯", "诸葛", "尉迟", "公羊", "赫连", "澹台",
					"皇甫", "宗政", "濮阳", "公冶", "太叔", "申屠", "公孙", "慕容",
					"钟离", "长孙", "司徒", "鲜于", "司空", "亓官", "司寇", "仉督",
					"子车", "颛孙", "端木", "巫马", "公西", "漆雕", "乐正", "壤驷",
					"公良", "拓跋", "夹谷", "宰父", "谷梁", "段干", "百里", "呼延",
					"东郭", "南门", "羊舌", "微生", "左丘", "东门", "西门", "第五",
					"言福", "刘付", "相里", "子书", "即墨", "达奚", "褚师", "况后",
					"梁丘", "东宫", "仲长", "屈突", "尔朱", "纳兰", "慕容", "尉迟",
					"可频", "纥豆陵", "宿勤", "阿跌", "斛律", "叱吕", "贺若", "秃发",
					"乞伏", "厍狄", "乌古论", "古里", "夹谷", "蒲察", "女奚烈", "兀颜",
					"斯陈", "孙伯", "归海", "后氏", "有氏", "琴氏", "冶氏", "厉氏"
				)

                # 提取姓和名，支持复姓
                if ($cnName.Length -ge 2 -and $doubleSurnames -contains $cnName.Substring(0, 2)) {
                    $surname = $cnName.Substring(0, 2)
                    $givenName = if ($cnName.Length -gt 2) { $cnName.Substring(2) } else { "" }
                } else {
                    $surname = $cnName.Substring(0, 1)
                    $givenName = if ($cnName.Length -gt 1) { $cnName.Substring(1) } else { $cnName }
                }

                # 创建安全密码
                $securePwd = ConvertTo-SecureString $password -AsPlainText -Force -ErrorAction Stop

                # 准备用户参数
                $userParams = @{
                    SamAccountName        = $username
                    UserPrincipalName     = "$username@$domainDNSRoot"
                    Name                  = $username
                    DisplayName           = $cnName
                    Surname               = $surname
                    GivenName             = $givenName
                    Description           = $description
                    EmailAddress          = $email
					OfficePhone           = $phone
                    Path                  = $NameOU
                    AccountPassword       = $securePwd
                    Enabled               = $true
                    ChangePasswordAtLogon = $true
                    ErrorAction           = "Stop"
                }

<#
                # 设置账号过期时间
                if (-not $neverExpire -and $expiryDate -gt (Get-Date)) {
                    $userParams.AccountExpirationDate = $expiryDate
                }
#>
                if (-not $neverExpire) {
                    $userParams.AccountExpirationDate = $expiryDate
                }


                # 创建用户 - 新用户会自动加入Domain Users组
                New-ADUser @userParams

                # 获取创建的用户信息
                $newUser = Get-ADUser -Identity $username -Properties DistinguishedName -ErrorAction Stop
                
                return @{ 
                    Success = $true 
                    DN = $newUser.DistinguishedName 
                }
            }
            ArgumentList = @($cnName, $username, $email, $phone, $currentOU, $description, $password, $expiryDate, $domainDNSRoot, $neverExpire)
            ErrorAction = "Stop"
        }

        $remoteResult = Invoke-Command @remoteParams

        [System.Windows.Forms.MessageBox]::Show("账号创建成功！`n账号：$username`n姓名：$cnName`n创建路径：$($remoteResult.DN)", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
        ClearInputFields  # 来自Helpers.ps1
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "创建失败: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("账号[$username]创建失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}




function ModifyUserAccount {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请选择需要修改的用户", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 获取选中账号信息
    $selectedRow = $script:userDataGridView.SelectedRows[0]
    $oldUsername = $null
    $oldDisplayName = $null

    if ($selectedRow.DataBoundItem -ne $null) {
        $userData = $selectedRow.DataBoundItem
        $oldUsername = $userData.SamAccountName.ToString().Trim()
        $oldDisplayName = $userData.DisplayName.ToString().Trim()
    }
    if ([string]::IsNullOrEmpty($oldUsername) -and $selectedRow.Cells["SamAccountName"] -ne $null) {
        $oldUsername = $selectedRow.Cells["SamAccountName"].Value.ToString().Trim()
    }

    if ([string]::IsNullOrEmpty($oldUsername)) {
        [System.Windows.Forms.MessageBox]::Show("未找到原账号信息", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 获取新输入的信息
    $newDisplayName = $script:textCnName.Text.Trim()
    $newEmail = $script:textEmail.Text.Trim()
	$newPhone = $script:textPhone.Text.Trim()
    $newDescription = $script:textDescription.Text.Trim()
    $neverExpire = $script:chkNeverExpire.Checked
    $expiryDate = $script:dateExpiry.Value.AddDays(1)

    # 验证是否有修改
	$oldPhone = if ($selectedRow.Cells["OfficePhone"] -ne $null) { $selectedRow.Cells["OfficePhone"].Value } else { "" }
    if ($newDisplayName -eq $oldDisplayName -and 
        $newEmail -eq $selectedRow.Cells["EmailAddress"].Value -and
        $newPhone -eq $oldPhone -and		
        $newDescription -eq $selectedRow.Cells["Description"].Value) {
        [System.Windows.Forms.MessageBox]::Show("未检测到任何修改", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    try {
        $script:connectionStatus = "正在远程修改账号[$oldUsername]信息..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 远程执行修改
        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($username, $displayName, $email, $phone, $description, $neverExpire, $expiryDate)
            Import-Module ActiveDirectory -ErrorAction Stop

            $updateParams = @{
                Identity = $username
                ErrorAction = "Stop"
            }

			#if (-not [string]::IsNullOrEmpty($displayName)){$updateParams.DisplayName = $displayName} 
			
            if ([string]::IsNullOrEmpty($email)) { 
				$updateParams.EmailAddress = $Null
				} elseif (-not [string]::IsNullOrEmpty($email)){
					$updateParams.EmailAddress = $email 
				}
			if ([string]::IsNullOrEmpty($phone)) {
				$updateParams.OfficePhone = $Null
				} elseif (-not [string]::IsNullOrEmpty($phone)){
					$updateParams.OfficePhone = $phone
				}
            if ([string]::IsNullOrEmpty($description)) { 
				$updateParams.Description = $Null
				} elseif (-not [string]::IsNullOrEmpty($description )){
					$updateParams.Description = $description 
				}

<#            
            if ($neverExpire) {
                $updateParams.AccountExpirationDate = $null
            } elseif ($expiryDate -gt (Get-Date)) {
                $updateParams.AccountExpirationDate = $expiryDate
            }						
#>
            if ($neverExpire) {
                $updateParams.AccountExpirationDate = $null
            } else {
                $updateParams.AccountExpirationDate = $expiryDate
            }	

            Set-ADUser @updateParams
            return $true
        } -ArgumentList $oldUsername, $newDisplayName, $newEmail, $newPhone, $newDescription, $neverExpire, $expiryDate -ErrorAction Stop

        [System.Windows.Forms.MessageBox]::Show("账号[$oldUsername]信息修改成功", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
        $script:connectionStatus = "已连接到域控: $($script:comboDomain.SelectedItem.Name)（远程执行）"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "修改失败: $errorMsg"
		Write-Error $errorMsg
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("账号[$oldUsername]信息修改失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}
    



function ChangeUserPassword {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请选择用户", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    # 获取选中账号
    $selectedRow = $script:userDataGridView.SelectedRows[0]
    $username = $selectedRow.Cells["SamAccountName"].Value.ToString().Trim()
    
    $password = $script:textNewPassword.Text
    $confirm = $script:textConfirmPassword.Text
    
    if (-not $password) {
        [System.Windows.Forms.MessageBox]::Show("请输入新密码", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($password -ne $confirm) {
        [System.Windows.Forms.MessageBox]::Show("两次密码不一致", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($password.Length -lt 8 -or $password -notmatch '[A-Z]' -or $password -notmatch '[a-z]' -or $password -notmatch '[0-9]' -or $password -notmatch '[^a-zA-Z0-9]') {
        [System.Windows.Forms.MessageBox]::Show("密码需≥8位，包含大、小写字母、数字和特殊字符", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    try {
        $script:connectionStatus = "正在远程修改密码..."
        UpdateStatusBar
        $script:mainForm.Refresh()
        
        # 远程执行密码修改
        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($username, $password)
            Import-Module ActiveDirectory -ErrorAction Stop
            $securePwd = ConvertTo-SecureString $password -AsPlainText -Force
            Set-ADAccountPassword -Identity $username -NewPassword $securePwd -Reset -ErrorAction Stop
            Set-ADUser -Identity $username -ChangePasswordAtLogon $true -ErrorAction Stop
        } -ArgumentList $username, $password -ErrorAction Stop
        
        [System.Windows.Forms.MessageBox]::Show("密码修改成功", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
        $script:textNewPassword.Text = ""
        $script:textConfirmPassword.Text = ""
    }
    catch {
        $script:connectionStatus = "修改失败: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("修改失败: $($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function UnlockUserAccount {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请选择用户", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    # 获取选中账号信息
    $selectedRow = $script:userDataGridView.SelectedRows[0]
    $username = $null
    $isLocked = $false

    if ($selectedRow.DataBoundItem -ne $null) {
        $userData = $selectedRow.DataBoundItem
        if ($userData.PSObject.Properties['SamAccountName']) {
            $username = $userData.SamAccountName.ToString().Trim()
        }
    }
    if ([string]::IsNullOrEmpty($username) -and $selectedRow.Cells["SamAccountName"] -ne $null) {
        $username = $selectedRow.Cells["SamAccountName"].Value.ToString().Trim()
    }
    
    if ([string]::IsNullOrEmpty($username)) {
        [System.Windows.Forms.MessageBox]::Show("未找到账号信息，请重试", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 远程查询账号锁定状态
    try {
        $script:connectionStatus = "正在查询账号锁定状态..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        $lockStatus = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($targetUser)
            Import-Module ActiveDirectory -ErrorAction Stop
            $user = Get-ADUser -Identity $targetUser -Properties LockedOut -ErrorAction Stop
            return $user.LockedOut
        } -ArgumentList $username -ErrorAction Stop

        $isLocked = [bool]$lockStatus
    }
    catch {
        $script:connectionStatus = "查询锁定状态失败: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("查询账号[$username]状态失败：`n$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 验证是否需要解锁
    if (-not $isLocked) {
        [System.Windows.Forms.MessageBox]::Show("账号[$username]未锁定，无需操作", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 确认解锁操作
    if ([System.Windows.Forms.MessageBox]::Show("确定解锁账号 [$username] 吗？", "确认解锁", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne 'Yes') {
        return
    }

    # 远程执行解锁
    try {
        $script:connectionStatus = "正在远程解锁账号[$username]..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($targetUser)
            Import-Module ActiveDirectory -ErrorAction Stop
            $user = Get-ADUser -Identity $targetUser -ErrorAction Stop
            Unlock-ADAccount -Identity $user.DistinguishedName -ErrorAction Stop
            $updatedUser = Get-ADUser -Identity $user.DistinguishedName -Properties LockedOut -ErrorAction Stop
            if ($updatedUser.LockedOut) {
                throw "解锁后账号仍处于锁定状态，可能是域控同步延迟"
            }
        } -ArgumentList $username -ErrorAction Stop

        [System.Windows.Forms.MessageBox]::Show("账号[$username]解锁成功", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
        $script:connectionStatus = "已连接到域控: $($script:comboDomain.SelectedItem.Name)（远程执行）"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "解锁失败: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("账号[$username]解锁失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function RenameUserAccount {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请选择需要重命名的用户", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 提取原账号信息
    $selectedRow = $script:userDataGridView.SelectedRows[0]
    $oldUsername = $null
    $oldDisplayName = $null

    if ($selectedRow.DataBoundItem -ne $null) {
        $userData = $selectedRow.DataBoundItem
        $oldUsername = $userData.SamAccountName.ToString().Trim()
        $oldDisplayName = $userData.DisplayName.ToString().Trim()
    }
    if ([string]::IsNullOrEmpty($oldUsername) -and $selectedRow.Cells["SamAccountName"] -ne $null) {
        $oldUsername = $selectedRow.Cells["SamAccountName"].Value.ToString().Trim()
    }
    if ([string]::IsNullOrEmpty($oldDisplayName) -and $selectedRow.Cells["DisplayName"] -ne $null) {
        $oldDisplayName = $selectedRow.Cells["DisplayName"].Value.ToString().Trim()
    }

    if ([string]::IsNullOrEmpty($oldUsername)) {
        [System.Windows.Forms.MessageBox]::Show("未找到原账号信息", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 获取新账号信息
    $newUsername = $script:textPinyin.Text.Trim()
    $newDisplayName = $script:textCnName.Text.Trim()

    if ([string]::IsNullOrEmpty($newUsername) -or [string]::IsNullOrEmpty($newDisplayName)) {
        [System.Windows.Forms.MessageBox]::Show("请输入新账号和新姓名", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($newUsername -eq $oldUsername -and $newDisplayName -eq $oldDisplayName) {
        [System.Windows.Forms.MessageBox]::Show("新信息与原信息一致，无需修改", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 远程检查新账号是否已存在
    try {
        $script:connectionStatus = "正在检查新账号可用性..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        $exists = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($newUser)
            Import-Module ActiveDirectory -ErrorAction Stop
            $user = Get-ADUser -Filter "SamAccountName -eq '$newUser'" -ErrorAction SilentlyContinue
            return $null -ne $user
        } -ArgumentList $newUsername -ErrorAction Stop

        if ($exists) {
            [System.Windows.Forms.MessageBox]::Show("新账号[$newUsername]已存在，请更换", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }
    catch {
        $script:connectionStatus = "检查账号失败: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("检查新账号可用性失败：`n$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 确认重命名操作
    $confirmMsg = "确定重命名账号？`n原信息：$oldDisplayName（$oldUsername）`n新信息：$newDisplayName（$newUsername）"
    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, "确认重命名", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne 'Yes') {
        return
    }

    # 远程执行重命名
    try {
        $script:connectionStatus = "正在远程重命名账号[$oldUsername]..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($oldUser, $newUser, $newDisplayName, $domainDNSRoot)
            Import-Module ActiveDirectory -ErrorAction Stop

            $user = Get-ADUser -Identity $oldUser -Properties DistinguishedName -ErrorAction Stop
            $userDN = $user.DistinguishedName

            # 修改SamAccountName和UPN
            Set-ADUser -Identity $userDN `
                       -SamAccountName $newUser `
                       -UserPrincipalName "$newUser@$domainDNSRoot" `
                       -ErrorAction Stop

            # 修改DisplayName
            Set-ADUser -Identity $newUser `
                       -DisplayName $newDisplayName `
                       -GivenName $newDisplayName `
                       -ErrorAction Stop

            # 验证修改结果
            $updatedUser = Get-ADUser -Identity $newUser -Properties DisplayName, UserPrincipalName -ErrorAction Stop
            if ($updatedUser.SamAccountName -ne $newUser -or $updatedUser.DisplayName -ne $newDisplayName) {
                throw "重命名后信息不匹配，修改未完全生效"
            }
        } -ArgumentList $oldUsername, $newUsername, $newDisplayName, $script:domainContext.DomainInfo.DNSRoot -ErrorAction Stop

        [System.Windows.Forms.MessageBox]::Show("账号重命名成功！`n新账号：$newUsername`n新姓名：$newDisplayName", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
        ClearInputFields
        $script:connectionStatus = "已连接到域控: $($script:comboDomain.SelectedItem.Name)（远程执行）"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "重命名失败: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("账号重命名失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}



function DeleteUserAccount {
    # 1. 基础校验：是否连接域控
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    # 获取所有选中行
    $selectedRows = $script:userDataGridView.SelectedRows
    if ($selectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请通过Ctrl键多选需要删除的用户", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }


    # 2. 提取所有选中行的用户信息（增加空值检查）
    $selectedUsers = @()
    foreach ($row in $selectedRows) {
        $userInfo = [PSCustomObject]@{
            DisplayName        = ""  # 默认为空字符串避免null
            SamAccountName     = ""
            IsValid            = $false
            DistinguishedName  = $null
        }

        # 从DataBoundItem提取（优先方式）
        if ($row.DataBoundItem -ne $null) {
            $userData = $row.DataBoundItem
            
            # 安全处理DisplayName（避免null调用方法）
            if ($null -ne $userData.DisplayName) {
                $userInfo.DisplayName = $userData.DisplayName.ToString().Trim()
            }
            
            # 安全处理SamAccountName（账号是关键信息，必须校验）
            if ($null -ne $userData.SamAccountName) {
                $userInfo.SamAccountName = $userData.SamAccountName.ToString().Trim()
            }
        }
        # 从单元格直接提取（备用方式）
        else {
            # 处理DisplayName单元格
            if ($row.Cells["DisplayName"] -ne $null -and $null -ne $row.Cells["DisplayName"].Value) {
                $userInfo.DisplayName = $row.Cells["DisplayName"].Value.ToString().Trim()
            }
            
            # 处理SamAccountName单元格（关键信息）
            if ($row.Cells["SamAccountName"] -ne $null -and $null -ne $row.Cells["SamAccountName"].Value) {
                $userInfo.SamAccountName = $row.Cells["SamAccountName"].Value.ToString().Trim()
            }
        }

        # 只保留有有效账号的用户
        if (-not [string]::IsNullOrEmpty($userInfo.SamAccountName)) {
            $selectedUsers += $userInfo
        }
    }

    # 校验有效用户数量
    if ($selectedUsers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("选中的行中未找到有效账号信息，请重新选择", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }


    # 3. 批量验证账号在域控中的有效性
    $script:connectionStatus = "正在验证 $($selectedUsers.Count) 个账号的有效性..."
    UpdateStatusBar
    $script:mainForm.Refresh()

    $validationFailed = $false
    foreach ($user in $selectedUsers) {
        try {
            $userDN = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($targetUser)
                Import-Module ActiveDirectory -ErrorAction Stop
                $adUser = Get-ADUser -Identity $targetUser -Properties DistinguishedName -ErrorAction Stop
                return $adUser.DistinguishedName
            } -ArgumentList $user.SamAccountName -ErrorAction Stop

            $user.IsValid = $true
            $user.DistinguishedName = $userDN
        }
        catch {
            $validationFailed = $true
            $errorMsg = $_.Exception.Message
            $script:connectionStatus = "账号验证失败: $($user.SamAccountName)"
            UpdateStatusBar
            [System.Windows.Forms.MessageBox]::Show(
                "账号 [$($user.SamAccountName)] 验证失败：`n$errorMsg", 
                "验证错误", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            break
        }
    }

    if ($validationFailed) {
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        return
    }


    # 4. 批量删除二次确认
    $confirmMsg = "确定永久删除以下 $($selectedUsers.Count) 个账号？`n`n"
    $confirmMsg += "序号 | 姓名 | 账号`n"
    $confirmMsg += "------------------------`n"
    for ($i = 0; $i -lt $selectedUsers.Count; $i++) {
        $user = $selectedUsers[$i]
        $confirmMsg += "$($i+1).   | $($user.DisplayName) | $($user.SamAccountName)`n"
    }
    $confirmMsg += "`n此操作不可恢复，删除后将无法恢复数据！"

    if ([System.Windows.Forms.MessageBox]::Show(
        $confirmMsg, 
        "批量删除警告", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) -ne 'Yes') {
        return
    }


    # 5. 执行批量删除
    $script:connectionStatus = "正在执行批量删除..."
    UpdateStatusBar
    $script:mainForm.Refresh()

    $deleteResults = [PSCustomObject]@{
        SuccessCount = 0
        FailedCount  = 0
        FailedUsers  = @()
    }

    foreach ($user in $selectedUsers) {
        try {
            Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($targetDN)
                Import-Module ActiveDirectory -ErrorAction Stop
                # 二次校验账号存在性
                $adUser = Get-ADUser -Filter "DistinguishedName -eq '$targetDN'" -ErrorAction Stop
                Remove-ADUser -Identity $targetDN -Confirm:$false -ErrorAction Stop
                # 验证删除结果
                $remaining = Get-ADUser -Filter "DistinguishedName -eq '$targetDN'" -ErrorAction SilentlyContinue
                if ($remaining) { throw "删除后仍能查询到账号（可能是AD缓存）" }
            } -ArgumentList $user.DistinguishedName -ErrorAction Stop

            $deleteResults.SuccessCount++
        }
        catch {
            $deleteResults.FailedCount++
            $deleteResults.FailedUsers += [PSCustomObject]@{
                Account  = $user.SamAccountName
                Name     = $user.DisplayName
                ErrorMsg = $_.Exception.Message
            }
        }
    }


    # 6. 显示删除结果
    $resultMsg = "删除完成！`n`n"
    $resultMsg += "成功删除：$($deleteResults.SuccessCount) 个账号`n"
    $resultMsg += "删除失败：$($deleteResults.FailedCount) 个账号`n"

    if ($deleteResults.FailedCount -gt 0) {
        $resultMsg += "`n失败详情：`n"
        foreach ($failed in $deleteResults.FailedUsers) {
            $resultMsg += "- $($failed.Account)（$($failed.Name)）：$($failed.ErrorMsg)`n"
        }
        [System.Windows.Forms.MessageBox]::Show($resultMsg, "操作结果", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $script:connectionStatus = "批量删除完成（部分失败）"
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::Orange
    }
    else {
        [System.Windows.Forms.MessageBox]::Show($resultMsg, "操作成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $script:connectionStatus = "批量删除成功"
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }


    # 7. 刷新界面
    $script:userDataGridView.ClearSelection()
    LoadUserList  # 刷新用户列表
    UpdateStatusBar
}







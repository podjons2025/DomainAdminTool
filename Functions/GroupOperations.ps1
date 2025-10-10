<# 
组相关核心操作 
#>

function LoadGroupList {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    try {
        $script:connectionStatus = "正在加载 OU: $($script:currentOU) 下的组..."
        UpdateStatusBar
        $script:mainForm.Refresh()
        
        $script:allGroups.Clear()
        $script:filteredGroups.Clear()
        
        # 远程加载组（逻辑不变）
        $remoteGroups = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($searchBase, $allUsersOU)
            Import-Module ActiveDirectory -ErrorAction Stop
            if ($allUsersOU) {
                Get-ADGroup -Filter * `
                    -Properties Name, SamAccountName, Description `
                    | Select-Object Name, SamAccountName, Description				
            } else {
                Get-ADGroup -Filter * -SearchBase $searchBase `
                    -Properties Name, SamAccountName, Description `
                    | Select-Object Name, SamAccountName, Description
            }
        } -ArgumentList $script:currentOU, $script:allUsersOU -ErrorAction Stop
        
        # 填充数据（逻辑不变）
        $remoteGroups | ForEach-Object {
            $null = $script:allGroups.Add($_)
            $null = $script:filteredGroups.Add($_)
        }
        
        # ---------------------- 关键：默认全显配置 ----------------------
        $script:groupDefaultShowAll = $true  # 切换OU后强制默认全显
        $script:currentGroupPage = 1  # 重置当前页码为1
        # 全显时总页数=1（用总数据量作为pageSize）
        $script:totalGroupPages = Get-TotalPages -totalCount $script:filteredGroups.Count -pageSize $script:filteredGroups.Count  
        
        # 1. 绑定全量数据到DataGridView
        $script:groupDataGridView.DataSource = $null
        $script:groupDataGridView.DataSource = $script:filteredGroups
        
        # 2. 同步分页控件状态
        $script:lblGroupPageInfo.Text = "第 $script:currentGroupPage 页 / 共 $script:totalGroupPages 页（总计 $($script:filteredGroups.Count) 条）"
        $script:btnGroupPrev.Enabled = $false
        $script:btnGroupNext.Enabled = $false
        $script:txtGroupJumpPage.Text = "1"
        $script:groupPaginationPanel.Visible = $true
        # ----------------------------------------------------------------
        
		# 首次加载后初始化动态分页大小
		Update-DynamicGroupPageSize
		
        # 更新状态（逻辑不变）
        $script:groupCountStatus = $script:allGroups.Count
        $script:connectionStatus = "已加载 OU: $($script:currentOU) 下的 $($script:groupCountStatus) 个组"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "加载组失败: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("组列表加载失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


function CreateNewGroup {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 获取输入信息
    $groupName = $script:textGroupName.Text.Trim()
    $groupSam = $script:textGroupSamAccount.Text.Trim()
    $groupDesc = $script:textGroupDescription.Text.Trim()

    if ([string]::IsNullOrEmpty($groupName) -or [string]::IsNullOrEmpty($groupSam)) {
        [System.Windows.Forms.MessageBox]::Show("组名称和组账号为必填项", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 远程检查组是否已存在
    try {
        $script:connectionStatus = "正在检查组可用性..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        $exists = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($groupSamAccount, $NameOU)
            Import-Module ActiveDirectory -ErrorAction Stop
            $group = Get-ADGroup -Filter "SamAccountName -eq '$groupSamAccount'" -ErrorAction SilentlyContinue
            return $null -ne $group
        } -ArgumentList $groupSam , $script:currentOU -ErrorAction Stop

        if ($exists) {
            [System.Windows.Forms.MessageBox]::Show("组账号[$groupSam]已存在，请更换", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }
    catch {
        $script:connectionStatus = "检查组失败: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("检查组可用性失败：`n$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 远程创建组
    try {
        $script:connectionStatus = "正在远程创建组[$groupName]..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($name, $sam, $desc, $domainPartition)
            Import-Module ActiveDirectory -ErrorAction Stop

            $groupParams = @{
                Name            = $name
                SamAccountName  = $sam
                Description     = $desc
                GroupCategory   = "Security"
				GroupScope      = "Global"
				Path            = $NameOU			
}

            New-ADGroup @groupParams
            $newGroup = Get-ADGroup -Identity $sam -Properties Description -ErrorAction Stop
            if ($newGroup.Name -ne $name -or $newGroup.Description -ne $desc) {
                throw "组创建成功，但属性不匹配"
            }
        } -ArgumentList $groupName, $groupSam, $groupDesc, $script:domainContext.DomainInfo.DefaultPartition -ErrorAction Stop

        [System.Windows.Forms.MessageBox]::Show("组[$groupName]创建成功", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadGroupList
        ClearGroupInputFields  # 来自Helpers.ps1
        $script:connectionStatus = "已连接到域控: $($script:comboDomain.SelectedItem.Name)（远程执行）"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "创建组失败: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("组创建失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


function AddUserToGroup {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 支持多选用户：检查是否选中至少1个用户（兼容单/多选）
    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请先选择需要加入组的用户（支持Ctrl多选）", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 检查是否选中目标组（单个组）
    if ($script:groupDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请先选择目标组（仅支持单个组）", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $selectedGroup = $script:groupDataGridView.SelectedRows[0].DataBoundItem
    if (-not $selectedGroup) {
        [System.Windows.Forms.MessageBox]::Show("选中组数据异常，请重新选择", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 提取所有选中用户的核心信息（SamAccountName + 显示名）
    $selectedUsers = @()
    foreach ($userRow in $script:userDataGridView.SelectedRows) {
        # 提取用户SamAccountName（AD操作唯一标识，必须非空）
        if ($userRow.Cells["SamAccountName"].Value -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("选中用户行数据异常（账号为空），请重新选择", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $username = $userRow.Cells["SamAccountName"].Value.ToString().Trim()

        # 提取用户显示名（用于提示，为空则用账号代替）
        $userDisplay = if ($userRow.Cells["DisplayName"].Value -ne $null) {
            $userRow.Cells["DisplayName"].Value.ToString().Trim()
        } else {
            $username
        }

        # 加入待处理用户列表
        $selectedUsers += [PSCustomObject]@{
            SamAccountName = $username
            DisplayName    = $userDisplay
        }
    }

    # 提取目标组信息
    if ($selectedGroup.SamAccountName -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("选中组的账号信息为空，请重新选择", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $groupSam = $selectedGroup.SamAccountName.ToString().Trim()
    $groupName = if ($selectedGroup.Name -ne $null) {
        $selectedGroup.Name.ToString().Trim()
    } else {
        $groupSam
    }

    # 批量检查用户是否已在组中
    $existingUsers = @()  # 已在组中的用户
    $validUsers = @()     # 待添加的有效用户
    try {
        $script:connectionStatus = "正在检查 $($selectedUsers.Count) 个用户的组成员关系..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 远程批量获取组内所有成员（减少AD调用次数，提升性能）
        $groupMembers = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($targetGroup)
            Import-Module ActiveDirectory -ErrorAction Stop
            $members = Get-ADGroupMember -Identity $targetGroup -Recursive -ErrorAction Stop
            return $members.SamAccountName  # 仅返回Sam账号，减少数据传输
        } -ArgumentList $groupSam -ErrorAction Stop

        # 对比筛选：已存在的用户 vs 待添加的用户
        foreach ($user in $selectedUsers) {
            if ($groupMembers -contains $user.SamAccountName) {
                $existingUsers += $user
            } else {
                $validUsers += $user
            }
        }

        # 提示已存在的用户（不中断操作，仅告知）
        if ($existingUsers.Count -gt 0) {
            $existingNames = @()
            foreach ($user in $existingUsers) {
                $existingNames += "$($user.DisplayName)（$($user.SamAccountName)）"
            }
            [System.Windows.Forms.MessageBox]::Show("以下用户已在组[$groupName]中，无需重复添加：`n`n$($existingNames -join "`n")", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }

        # 若所有用户都已存在，直接退出
        if ($validUsers.Count -eq 0) {
            $script:connectionStatus = "已连接到域控: $($script:comboDomain.SelectedItem.Name)（远程执行）"
            UpdateStatusBar
            return
        }
    }
    catch {
        $script:connectionStatus = "检查成员关系失败: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("检查用户组关系失败：`n$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 批量确认添加操作（区分单/多用户提示文案）
    $validUserNames = @()
    foreach ($user in $validUsers) {
        $validUserNames += "$($user.DisplayName)（$($user.SamAccountName)）"
    }
    
    $confirmTitle = if ($validUsers.Count -eq 1) { "确认添加用户到组" } else { "确认批量添加用户到组" }
    $confirmMsg = if ($validUsers.Count -eq 1) {
        "确定将用户`n`n$($validUserNames -join "`n")`n`n加入组[$groupName（$groupSam）]吗？"
    } else {
        "共选中 $($validUsers.Count) 个用户，确定批量加入组[$groupName（$groupSam）]吗？`n`n待添加用户：`n$($validUserNames -join "`n")"
    }

    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, $confirmTitle, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne 'Yes') {
        return
    }

    # 批量远程添加用户到组（一次传递所有用户Sam账号，减少AD调用）
    try {
        $script:connectionStatus = "正在远程添加 $($validUsers.Count) 个用户到组[$groupName]..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 提取待添加用户的Sam账号列表（AD批量操作需此格式）
        $validUserSams = $validUsers.SamAccountName

        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($users, $group)
            Import-Module ActiveDirectory -ErrorAction Stop
            # 批量添加（Add-ADGroupMember支持多成员参数）
            Add-ADGroupMember -Identity $group -Members $users -ErrorAction Stop
            
            # 验证：确保所有用户都已加入
            $updatedMembers = Get-ADGroupMember -Identity $group -Recursive -ErrorAction Stop
            $missingUsers = $users | Where-Object { $_ -notin $updatedMembers.SamAccountName }
            if ($missingUsers.Count -gt 0) {
                throw "部分用户添加失败：$($missingUsers -join ', ')"
            }
        } -ArgumentList $validUserSams, $groupSam -ErrorAction Stop

        # 批量成功提示
        $successMsg = if ($validUsers.Count -eq 1) {
            "用户[$($validUsers[0].DisplayName)]已成功加入组[$groupName]"
        } else {
            $successUserList = @()
            foreach ($user in $validUsers) {
                $successUserList += "$($user.DisplayName)（$($user.SamAccountName)）"
            }
            "已成功添加 $($validUsers.Count) 个用户到组[$groupName]：`n`n$($successUserList -join "`n")"
        }
        [System.Windows.Forms.MessageBox]::Show($successMsg, "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        LoadUserList  # 刷新用户列表（更新用户所属组信息）
        $script:connectionStatus = "已连接到域控: $($script:comboDomain.SelectedItem.Name)（远程执行）"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "添加失败: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("添加用户到组失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


function ModifyGroup {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    if ($script:groupDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请选择组", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    # 1. 获取选中组的基础信息（从DataGridView）
    $groupRow = $script:groupDataGridView.SelectedRows[0]
    $groupData = $groupRow.DataBoundItem
    $originalSam = $groupData.SamAccountName.ToString().Trim()
    
    if ([string]::IsNullOrEmpty($originalSam)) {
        [System.Windows.Forms.MessageBox]::Show("选中组的账号信息为空，请重新选择", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 2. 远程获取组的完整属性（含DN、CN、DisplayName、成员数量）
    # 新增：$hasMembers 标记组是否存在组员；$checkNestedMembers 控制是否检测嵌套成员
    $hasMembers = $false
    $checkNestedMembers = $false  # 按需调整：$true=包含嵌套成员，$false=仅直接成员
    try {
        $script:connectionStatus = "正在加载组详细信息..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        $originalGroup = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($targetSam, $checkNested)
            Import-Module ActiveDirectory -ErrorAction Stop
            # 1. 获取组基础属性 + Member属性（用于计算直接成员）
            $adGroup = Get-ADGroup -Identity $targetSam `
                        -Properties DistinguishedName, DisplayName, Description, Member `
                        -ErrorAction Stop
            
            # 2. 计算成员数量（区分直接/嵌套）
            if ($checkNested) {
                # 包含嵌套成员（递归获取）
                $allMembers = Get-ADGroupMember -Identity $adGroup -Recursive -ErrorAction Stop
                $memberCount = $allMembers.Count
            }
            else {
                # 仅直接成员（Member属性存储直接成员DN，空则为0）
                $memberCount = if ($adGroup.Member -and $adGroup.Member.Count -gt 0) { $adGroup.Member.Count } else { 0 }
            }

            # 3. 返回组属性 + 成员数量（用于后续判断是否需要刷新用户列表）
            return [PSCustomObject]@{
                Name              = $adGroup.Name
                DistinguishedName = $adGroup.DistinguishedName
                DisplayName       = $adGroup.DisplayName
                Description       = $adGroup.Description
                MemberCount       = $memberCount  # 新增：成员数量
            }
        } -ArgumentList $originalSam, $checkNestedMembers -ErrorAction Stop

        # 提取远程返回的关键属性（本地缓存）
        $originalCN = $originalGroup.Name          
        $originalDN = $originalGroup.DistinguishedName  
        $originalDisplayName = $originalGroup.DisplayName
        $originalDescription = $originalGroup.Description
        $originalMemberCount = $originalGroup.MemberCount  # 新增：记录原始成员数量
        
        # 新增：标记组是否存在组员（成员数 > 0 则为$true）
        if ($originalMemberCount -gt 0) {
            $hasMembers = $true
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "加载组信息失败: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("加载组详细信息失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # 3. 获取用户输入的新值并做本地格式校验（原逻辑不变）
    $newCN = $script:textGroupName.Text.Trim()          
    $newSam = $script:textGroupSamAccount.Text.Trim()   
    $newDisplayName = $newCN                            
    $newDesc = $script:textGroupDescription.Text.Trim() 

    # 3.1 必填项校验
    if (-not ($newCN -and $newSam)) {
        [System.Windows.Forms.MessageBox]::Show("组名称（CN）和组账号为必填项", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 3.2 CN格式校验（禁止LDAP特殊字符）
    if ($newCN -match '[,=\+<>;#\"\\]') {
        [System.Windows.Forms.MessageBox]::Show('组名称不能包含以下特殊字符：, = + < > ; # " \', "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 3.3 SamAccountName格式校验（AD内置限制）
    if ($newSam -match "[^\w\-]") {
        [System.Windows.Forms.MessageBox]::Show("组账号不能包含特殊字符（允许字母、数字、下划线、连字符）", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($newSam.Length -gt 20) {
        [System.Windows.Forms.MessageBox]::Show("组账号长度不能超过20个字符", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 4. 远程检查新SamAccountName的唯一性（若Sam有修改，原逻辑不变）
    if ($newSam -ne $originalSam) {
        try {
            $script:connectionStatus = "正在检查新组账号可用性..."
            UpdateStatusBar
            $script:mainForm.Refresh()

            $samExists = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($newSamAccount)
                Import-Module ActiveDirectory -ErrorAction Stop
                $existing = Get-ADGroup -Filter "SamAccountName -eq '$newSamAccount'" -ErrorAction SilentlyContinue
                return $null -ne $existing
            } -ArgumentList $newSam -ErrorAction Stop

            if ($samExists) {
                [System.Windows.Forms.MessageBox]::Show("新组账号[$newSam]已存在，请更换", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            $script:connectionStatus = "检查账号可用性失败: $errorMsg"
            UpdateStatusBar
            $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
            [System.Windows.Forms.MessageBox]::Show("检查新组账号失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }

    # 5. 检查是否有实际修改（无修改则直接返回，原逻辑不变）
    if ($newCN -eq $originalCN -and $newSam -eq $originalSam -and $newDisplayName -eq $originalDisplayName -and $newDesc -eq $originalDescription) {
        [System.Windows.Forms.MessageBox]::Show("未检测到任何修改", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 6. 确认修改操作（原逻辑不变，仅优化描述空值显示）
    $displayNewDesc = if ([string]::IsNullOrEmpty($newDesc)) { "无" } else { $newDesc }
    $confirmMsg = "确定修改组【$originalCN（$originalSam）】吗？`n"
    $confirmMsg += "注意：修改组名称（CN）会改变其在AD中的目录路径（DN）`n`n"
    $confirmMsg += "新名称（CN）：$newCN`n"
    $confirmMsg += "新账号（Sam）：$newSam`n"
    $confirmMsg += "新描述：$displayNewDesc"

    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, "确认修改组", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne 'Yes') {
        return
    }

    # 7. 远程执行修改操作（重命名+属性更新，原逻辑不变）
    try {
        $script:connectionStatus = "正在远程修改组【$originalCN】..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        $modifiedGroup = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($origSam, $origDN, $origCN, $newCN, $newSam, $newDisplayName, $newDesc)
            Import-Module ActiveDirectory -ErrorAction Stop

            # 若修改组名称（CN），先重命名AD对象
            if ($newCN -ne $origCN) {
                $newDN = $origDN -replace "^CN=$([regex]::Escape($origCN)),", "CN=$newCN,"
                Rename-ADObject -Identity $origDN `
                               -NewName $newCN `
                               -ErrorAction Stop
            }

            # 更新组属性（SamAccountName、DisplayName、Description）
            Set-ADGroup -Identity $origSam `
                        -SamAccountName $newSam `
                        -DisplayName $newDisplayName `
                        -Description $newDesc `
                        -ErrorAction Stop

            # 验证修改结果并返回最新属性
            $verifyIdentity = if ($newSam -eq $origSam) { $origSam } else { $newSam }
            return Get-ADGroup -Identity $verifyIdentity `
                               -Properties DistinguishedName, DisplayName, Description `
                               -ErrorAction Stop
        } -ArgumentList $originalSam, $originalDN, $originalCN, $newCN, $newSam, $newDisplayName, $newDesc -ErrorAction Stop

        # 8. 本地验证并提示成功（原逻辑不变）
        if ($modifiedGroup.Name -ne $newCN -or $modifiedGroup.SamAccountName -ne $newSam) {
            throw "远程修改执行成功，但返回的属性不匹配"
        }
        $displayModifiedDesc = if ([string]::IsNullOrEmpty($modifiedGroup.Description)) { "无" } else { $modifiedGroup.Description }
        [System.Windows.Forms.MessageBox]::Show("组修改成功`n`n" +
            "组名称（CN）：$($modifiedGroup.Name)`n" +
            "组账号（Sam）：$($modifiedGroup.SamAccountName)`n" +
            "显示名称：$($modifiedGroup.DisplayName)`n" +
            "描述：$displayModifiedDesc", 
            "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # 9. 新增：条件刷新用户列表（若组存在组员，则刷新用户的“所属组”信息）
        LoadGroupList  # 始终刷新组列表（展示修改后的组信息）
        if ($hasMembers) {
            LoadUserList   # 仅当组存在组员时，刷新用户列表（同步用户的所属组数据）
            $refreshTip = "（已同步刷新用户列表）"
        } else {
            $refreshTip = "（未刷新用户列表：组无组员）"
        }

        # 10. 更新状态栏（补充刷新状态提示）
        $script:connectionStatus = "已连接到域控: $($script:comboDomain.SelectedItem.Name)（远程执行） $refreshTip"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "修改组失败: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("组修改失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


function DeleteGroup {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 配置开关：是否检查嵌套成员（按需启用，大型组可能影响性能）
    $checkNestedMembers = $false  # $true=包含嵌套成员，$false=仅直接成员

    # 1. 检查是否选中组（支持Ctrl多选）
    if ($script:groupDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请选择需要删除的组", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 2. 提取选中组信息 + 精准检查成员数量（核心修复）
    $selectedGroups = @()
    $script:connectionStatus = "正在检查选中组的成员情况..."
    UpdateStatusBar
    $script:mainForm.Refresh()

    foreach ($row in $script:groupDataGridView.SelectedRows) {
        $group = $row.DataBoundItem
        if (-not $group) { continue }

        # 基础信息提取
        $groupSam = if ($group.SamAccountName -ne $null) { $group.SamAccountName.ToString().Trim() } else { "" }
        $groupName = if ($group.Name -ne $null) { $group.Name.ToString().Trim() } else { "未知组" }

        if ([string]::IsNullOrEmpty($groupSam)) {
            Write-Warning "跳过无效组（账号为空）：$groupName"
            continue
        }

        # 【核心修复】使用Get-ADGroup的Member属性检查成员（更可靠）
        try {
            $memberCount = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($samAccountName, $checkNested)
                Import-Module ActiveDirectory -ErrorAction Stop

                # 第一步：获取组对象（获取Member属性）
                $adGroup = Get-ADGroup -Identity $samAccountName -Properties Member -ErrorAction Stop

                # 第二步：计算成员数量（区分直接/嵌套）
                if ($checkNested) {
                    # 包含嵌套成员（递归获取所有用户/计算机）
                    $allMembers = Get-ADGroupMember -Identity $adGroup -Recursive -ErrorAction Stop
                    return $allMembers.Count
                }
                else {
                    # 仅直接成员（Member属性包含所有直接成员的DN）
                    if ($adGroup.Member -and $adGroup.Member.Count -gt 0) {
                        return $adGroup.Member.Count
                    }
                    else {
                        return 0
                    }
                }
            } -ArgumentList $groupSam, $checkNestedMembers -ErrorAction Stop
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            # 组不存在
            $memberCount = -2
            Write-Warning "检查组[$groupName]失败：组不存在"
        }
        catch [System.UnauthorizedAccessException] {
            # 权限不足
            $memberCount = -3
            Write-Warning "检查组[$groupName]成员失败：权限不足"
        }
        catch {
            # 其他错误
            $memberCount = -1
            Write-Warning "检查组[$groupName]成员失败：$($_.Exception.Message)"
        }

        # 加入选中列表（包含详细状态）
        $selectedGroups += [PSCustomObject]@{
            SamAccountName = $groupSam
            Name           = $groupName
            MemberCount    = $memberCount  # -3=权限不足；-2=组不存在；-1=其他错误；>=0=成员数
            CheckType      = if ($checkNestedMembers) { "包含嵌套" } else { "仅直接" }
        }
    }

    # 3. 过滤无效组并提示
    #$validGroups = $selectedGroups | Where-Object { $_.MemberCount -ge 0 }
    #$invalidGroups = $selectedGroups | Where-Object { $_.MemberCount -lt 0 }
	$validGroups = @($selectedGroups | Where-Object { $_.MemberCount -ge 0 })
	$invalidGroups = @($selectedGroups | Where-Object { $_.MemberCount -lt 0 })

    if ($invalidGroups.Count -gt 0) {
        $invalidMsg = "以下组无法删除：`n"
        foreach ($g in $invalidGroups) {
            switch ($g.MemberCount) {
                -3 { $invalidMsg += "- $($g.Name)：权限不足，无法访问成员信息`n" }
                -2 { $invalidMsg += "- $($g.Name)：组不存在`n" }
                default { $invalidMsg += "- $($g.Name)：成员检查失败`n" }
            }
        }
        [System.Windows.Forms.MessageBox]::Show($invalidMsg, "无法处理的组", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }

    if ($validGroups.Count -eq 0) {
        $script:connectionStatus = "已连接到域控: $($script:comboDomain.SelectedItem.Name)（远程执行）"
        UpdateStatusBar
        return
    }

    # 4. 生成确认提示（【重点修复】优化单选/多选判断，确保数量显示准确）
	$groupInfoLines = @()
	foreach ($g in $validGroups) {
		if ($g.MemberCount -gt 0) {
			$groupInfoLines += "$($g.Name)（$($g.SamAccountName)）- $($g.CheckType)成员：$($g.MemberCount)个"
		}
		else {
			$groupInfoLines += "$($g.Name)（$($g.SamAccountName)）- 无成员"
		}
	}
	$groupListText = $groupInfoLines -join "`n"

    # 【修复核心】明确区分单选/多选场景，避免变量空值
	if ($validGroups.Count -eq 1) {
		$confirmMsg = "确定永久删除以下组吗？`n`n$groupListText`n`n注意：将忽略成员直接删除！"
	}
	else {
		$confirmMsg = "共选中 $($validGroups.Count) 个有效组，确定永久删除吗？`n`n$groupListText`n`n注意：将忽略成员直接删除！"
	}

    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, "确认删除组", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -ne 'Yes') {
        $script:connectionStatus = "已连接到域控: $($script:comboDomain.SelectedItem.Name)（远程执行）"
        UpdateStatusBar
        return
    }

    # 5. 批量删除组 + 标记是否需要刷新用户列表
    $successCount = 0
    $failedGroups = @()
    $hasDeletedGroupWithMembers = $false
    $script:connectionStatus = "正在删除 $($validGroups.Count) 个组..."
    UpdateStatusBar
    $script:mainForm.Refresh()

    foreach ($g in $validGroups) {
        try {
            # 远程执行删除
            Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($groupSam)
                Import-Module ActiveDirectory -ErrorAction Stop
                # 强制删除（即使有成员）
                Remove-ADGroup -Identity $groupSam -Confirm:$false -ErrorAction Stop
                # 验证删除
                $exists = Get-ADGroup -Filter "SamAccountName -eq ""$groupSam""" -ErrorAction SilentlyContinue
                if ($exists) { throw "删除后仍可查询到组" }
            } -ArgumentList $g.SamAccountName -ErrorAction Stop

            $successCount++
            # 若删除的是有成员的组，标记刷新
            if ($g.MemberCount -gt 0) {
                $hasDeletedGroupWithMembers = $true
            }
        }
        catch {
            $failedGroups += [PSCustomObject]@{
                Name  = $g.Name
                Error = $_.Exception.Message
            }
        }
    }

    # 6. 显示删除结果
    $resultMsg = if ($successCount -eq $validGroups.Count) {
        if ($successCount -eq 1) {
            "组[$($validGroups[0].Name)]已永久删除"
        }
        else {
            $successList = $validGroups.Name -join "、"
            "已成功删除 $successCount 个组：`n$successList"
        }
    }
    else {
        $successNames = $validGroups | Where-Object { $_.SamAccountName -notin $failedGroups.SamAccountName } | Select-Object -ExpandProperty Name
        $successList = $successNames -join "、"
        $failedList = $failedGroups | ForEach-Object { "$($_.Name)（错误：$($_.Error)）" } -join "`n"
        "删除完成：`n成功：$successCount 个组（$successList）`n失败：$($failedGroups.Count) 个组：`n$failedList"
    }

    $msgIcon = if ($failedGroups.Count -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
    [System.Windows.Forms.MessageBox]::Show($resultMsg, "删除结果", [System.Windows.Forms.MessageBoxButtons]::OK, $msgIcon)

    # 7. 条件刷新列表
    $script:groupDataGridView.ClearSelection()
    LoadGroupList  # 始终刷新组列表

    if ($hasDeletedGroupWithMembers) {
        LoadUserList   # 仅删除有成员的组时刷新用户列表
        $refreshTip = "（已同步刷新用户列表）"
    }
    else {
        $refreshTip = "（未刷新用户列表：删除的均为无成员组）"
    }

    # 8. 更新状态栏
    $statusText = if ($failedGroups.Count -eq 0) {
        "已成功删除 $successCount 个组 $refreshTip"
    }
    else {
        "删除完成：成功 $successCount 个，失败 $($failedGroups.Count) 个 $refreshTip"
    }
    $script:connectionStatus = $statusText
    UpdateStatusBar
    $script:statusOutputLabel.ForeColor = if ($failedGroups.Count -eq 0) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkOrange }
}




function RemoveUserFromGroup {
    # 1. 检查域控连接
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 支持多选用户：检查选中用户数量
    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请先选择需要从组中移除的用户（支持Ctrl多选）", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 检查选中目标组（单个组）
    if ($script:groupDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请先选择要移除用户的目标组（仅支持单个组）", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $selectedGroup = $script:groupDataGridView.SelectedRows[0].DataBoundItem
    if (-not $selectedGroup) {
        [System.Windows.Forms.MessageBox]::Show("选中组数据异常，请重新选择", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 提取所有选中用户的核心信息
    $selectedUsers = @()
    foreach ($userRow in $script:userDataGridView.SelectedRows) {
        # 提取用户SamAccountName（必须非空）
        if ($userRow.Cells["SamAccountName"].Value -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("选中用户行数据异常（账号为空），请重新选择", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $username = $userRow.Cells["SamAccountName"].Value.ToString().Trim()

        # 提取用户显示名（为空则用账号代替）
        $userDisplay = if ($userRow.Cells["DisplayName"].Value -ne $null) {
            $userRow.Cells["DisplayName"].Value.ToString().Trim()
        } else {
            $username
        }

        $selectedUsers += [PSCustomObject]@{
            SamAccountName = $username
            DisplayName    = $userDisplay
        }
    }

    # 提取目标组信息
    if ($selectedGroup.SamAccountName -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("选中组的账号信息为空，请重新选择", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $groupSam = $selectedGroup.SamAccountName.ToString().Trim()
    $groupName = if ($selectedGroup.Name -ne $null) {
        $selectedGroup.Name.ToString().Trim()
    } else {
        $groupSam
    }

    # 批量检查用户是否在组中
    $nonMemberUsers = @()  # 不在组中的用户
    $validUsers = @()     # 待删除的有效用户
    try {
        $script:connectionStatus = "正在检查 $($selectedUsers.Count) 个用户的组成员关系..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 远程批量获取组内所有成员（减少AD调用）
        $groupMembers = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($targetGroup)
            Import-Module ActiveDirectory -ErrorAction Stop
            $members = Get-ADGroupMember -Identity $targetGroup -Recursive -ErrorAction Stop
            return $members.SamAccountName
        } -ArgumentList $groupSam -ErrorAction Stop

        # 对比筛选：不在组的用户 vs 待删除的用户
        foreach ($user in $selectedUsers) {
            if ($groupMembers -notcontains $user.SamAccountName) {
                $nonMemberUsers += $user
            } else {
                $validUsers += $user
            }
        }

        # 提示不在组的用户（不中断操作）
        if ($nonMemberUsers.Count -gt 0) {
            $nonMemberNames = @()
            foreach ($user in $nonMemberUsers) {
                $nonMemberNames += "$($user.DisplayName)（$($user.SamAccountName)）"
            }
            [System.Windows.Forms.MessageBox]::Show("以下用户不在组[$groupName]中，无需移除：`n`n$($nonMemberNames -join "`n")", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }

        # 若所有用户都不在组，直接退出
        if ($validUsers.Count -eq 0) {
            $script:connectionStatus = "已连接到域控: $($script:comboDomain.SelectedItem.Name)（远程执行）"
            UpdateStatusBar
            return
        }
    }
    catch {
        $script:connectionStatus = "检查成员关系失败: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("检查用户-组关系失败：`n$($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 批量确认删除操作
    $validUserNames = @()
    foreach ($user in $validUsers) {
        $validUserNames += "$($user.DisplayName)（$($user.SamAccountName)）"
    }
    
    $confirmTitle = if ($validUsers.Count -eq 1) { "确认移除用户" } else { "确认批量移除用户" }
    $confirmMsg = if ($validUsers.Count -eq 1) {
        "确定将用户`n`n$($validUserNames -join "`n")`n`n从组[$groupName（$groupSam）]中移除吗？`n`n移除后用户将失去该组的权限！"
    } else {
        "共选中 $($validUsers.Count) 个用户，确定批量从组[$groupName（$groupSam）]中移除吗？`n`n待移除用户：`n$($validUserNames -join "`n")`n`n移除后用户将失去该组的权限！"
    }

    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, $confirmTitle, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -ne 'Yes') {
        return
    }

    # 批量远程删除用户
    try {
        $script:connectionStatus = "正在远程移除 $($validUsers.Count) 个用户from组[$groupName]..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # 提取待删除用户的Sam账号列表
        $validUserSams = $validUsers.SamAccountName

        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($users, $group)
            Import-Module ActiveDirectory -ErrorAction Stop
            # 批量移除
            Remove-ADGroupMember -Identity $group -Members $users -Confirm:$false -ErrorAction Stop
            
            # 验证：确保所有用户都已移除
            $remainingMembers = Get-ADGroupMember -Identity $group -Recursive -ErrorAction Stop
            $remainingUsers = $users | Where-Object { $_ -in $remainingMembers.SamAccountName }
            if ($remainingUsers.Count -gt 0) {
                throw "部分用户移除失败：$($remainingUsers -join ', ')"
            }
        } -ArgumentList $validUserSams, $groupSam -ErrorAction Stop

        # 批量成功提示
        $successMsg = if ($validUsers.Count -eq 1) {
            "用户[$($validUsers[0].DisplayName)]已成功从组[$groupName]中移除"
        } else {
            $successUserList = @()
            foreach ($user in $validUsers) {
                $successUserList += "$($user.DisplayName)（$($user.SamAccountName)）"
            }
            "已成功从组[$groupName]中移除 $($validUsers.Count) 个用户：`n`n$($successUserList -join "`n")"
        }
        [System.Windows.Forms.MessageBox]::Show($successMsg, "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        LoadUserList  # 刷新用户列表（更新用户所属组信息）
        $script:connectionStatus = "已连接到域控: $($script:comboDomain.SelectedItem.Name)（远程执行）"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "移除用户失败: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("移除用户from组失败：`n$errorMsg", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

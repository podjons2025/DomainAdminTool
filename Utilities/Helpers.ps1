<# 
通用辅助函数 
#>

# 清空用户操作区输入框
function ClearInputFields {
    $script:textCnName.Text = ""
    $script:textPinyin.Text = ""
    $script:textEmail.Text = ""
	$script:textPhone.Text = ""
    $script:textDescription.Text = ""
    $script:textNewPassword.Text = ""
    $script:textConfirmPassword.Text = ""
    $script:chkNeverExpire.Checked = $true
    $script:dateExpiry.Enabled = $false
}

# 清空组操作区输入框
function ClearGroupInputFields {
    $script:textGroupName.Text = ""
    $script:textGroupSamAccount.Text = ""
    $script:textGroupDescription.Text = ""
    $script:originalGroupSamAccount = $null
	$script:groupDataGridView.ClearSelection()
}

# 筛选用户列表
function FilterUserList {
    param([string]$filterText)
    
    if ([string]::IsNullOrEmpty($filterText)) {
        $script:userDataGridView.DataSource = $script:allUsers
    } else {
        $lowerFilter = $filterText.ToLower()
        $filtered = @($script:allUsers | Where-Object {
            ( (-not [string]::IsNullOrEmpty($_.DisplayName)) -and $_.DisplayName.ToLower() -like "*$lowerFilter*" ) -or
            $_.SamAccountName.ToLower() -like "*$lowerFilter*" -or
            ( (-not [string]::IsNullOrEmpty($_.EmailAddress)) -and $_.EmailAddress.ToLower() -like "*$lowerFilter*" ) -or
			( (-not [string]::IsNullOrEmpty($_.TelePhone)) -and $_.TelePhone.ToLower() -like "*$lowerFilter*" ) -or
            ( (-not [string]::IsNullOrEmpty($_.Description)) -and $_.Description.ToLower() -like "*$lowerFilter*" ) -or
            ( (-not [string]::IsNullOrEmpty($_.MemberOf)) -and $_.MemberOf.ToLower() -like "*$lowerFilter*" )
        })
        $userList = New-Object System.Collections.ArrayList
        $userList.AddRange($filtered) | Out-Null  
        $script:userDataGridView.DataSource = $userList
        $script:userDataGridView.Refresh()
    }
}


# 筛选组列表
function FilterGroupList {
    param([string]$filterText)
    $script:groupDataGridView.DataSource = $null
	
    if ([string]::IsNullOrEmpty($filterText)) {
		$script:groupDataGridView.DataSource = $script:allGroups
    } else {
        $lowerFilter = $filterText.ToLower()
        $filtered = @($script:allGroups | Where-Object {
            $_.Name.ToLower() -like "*$lowerFilter*" -or
            $_.SamAccountName.ToLower() -like "*$lowerFilter*" -or
            ( (-not [string]::IsNullOrEmpty($_.Description)) -and $_.Description.ToLower() -like "*$lowerFilter*" )
        })
        $groupList = New-Object System.Collections.ArrayList
        $groupList.AddRange($filtered) | Out-Null  
        $script:groupDataGridView.DataSource = $groupList
        $script:groupDataGridView.Refresh()
    }
}

# 分页显示函数
function Show-UserPage {
    # 处理空数据（原逻辑不变）
    if ($script:filteredUsers.Count -eq 0) {
        $script:userDataGridView.DataSource = $null
        $script:lblUserPageInfo.Text = "第 0 页 / 共 0 页（总计 0 条）"
        $script:btnUserPrev.Enabled = $false
        $script:btnUserNext.Enabled = $false
        $script:txtUserJumpPage.Text = ""
        $script:userPaginationPanel.Visible = $false
        return
    }

    # ---------------------- 新增：判断是否默认全显 ----------------------
    if ($script:defaultShowAll) {
        # 全显模式：显示所有数据，分页控件禁用翻页
        $script:userDataGridView.DataSource = $null
        $script:userDataGridView.DataSource = $script:filteredUsers
        $script:totalUserPages = 1  # 全显时总页数=1
        $script:currentUserPage = 1  # 全显时页码=1
    }
    else {
        # 分页模式：按原逻辑截取当前页数据
        $startIndex = ($script:currentUserPage - 1) * $script:pageSize
        $endIndex = [math]::Min($startIndex + $script:pageSize, $script:filteredUsers.Count)
        $currentPageData = New-Object System.Collections.ArrayList
        for ($i = $startIndex; $i -lt $endIndex; $i++) {
            if ($null -ne $script:filteredUsers[$i]) {
                $null = $currentPageData.Add($script:filteredUsers[$i])
            }
        }
        # 绑定分页数据
        $script:userDataGridView.DataSource = $null
        $script:userDataGridView.DataSource = $currentPageData
        # 计算总页数（分页模式下用pageSize计算）
        $script:totalUserPages = Get-TotalPages -totalCount $script:filteredUsers.Count -pageSize $script:pageSize
    }

    # ---------------------- 更新分页控件状态 ----------------------
    $script:lblUserPageInfo.Text = "第 $script:currentUserPage 页 / 共 $script:totalUserPages 页（总计 $($script:filteredUsers.Count) 条）"
    $script:btnUserPrev.Enabled = ($script:currentUserPage -gt 1)  # 只有页码>1时可上一页
    $script:btnUserNext.Enabled = ($script:currentUserPage -lt $script:totalUserPages)  # 页码<总页数时可下一页
    $script:txtUserJumpPage.Text = $script:currentUserPage.ToString()  # 同步页码到跳转框
    $script:userPaginationPanel.Visible = $true  # 始终显示分页控件
}


function Show-GroupPage {
    # 处理空数据（原逻辑不变）
    if ($script:filteredGroups.Count -eq 0) {
        $script:groupDataGridView.DataSource = $null
        $script:lblGroupPageInfo.Text = "第 0 页 / 共 0 页（总计 0 条）"
        $script:btnGroupPrev.Enabled = $false
        $script:btnGroupNext.Enabled = $false
        $script:txtGroupJumpPage.Text = ""
        $script:groupPaginationPanel.Visible = $false
        return
    }

    # ---------------------- 新增：判断是否默认全显 ----------------------
    if ($script:groupDefaultShowAll) {
        # 全显模式：显示所有数据
        $script:groupDataGridView.DataSource = $null
        $script:groupDataGridView.DataSource = $script:filteredGroups
        $script:totalGroupPages = 1
        $script:currentGroupPage = 1
    }
    else {
        # 分页模式：截取当前页数据
        $startIndex = ($script:currentGroupPage - 1) * $script:pageSize
        $endIndex = [math]::Min($startIndex + $script:pageSize, $script:filteredGroups.Count)
        $currentPageData = New-Object System.Collections.ArrayList
        for ($i = $startIndex; $i -lt $endIndex; $i++) {
            if ($null -ne $script:filteredGroups[$i]) {
                $null = $currentPageData.Add($script:filteredGroups[$i])
            }
        }
        # 绑定分页数据
        $script:groupDataGridView.DataSource = $null
        $script:groupDataGridView.DataSource = $currentPageData
        # 计算总页数
        $script:totalGroupPages = Get-TotalPages -totalCount $script:filteredGroups.Count -pageSize $script:pageSize
    }

    # ---------------------- 更新分页控件状态 ----------------------
    $script:lblGroupPageInfo.Text = "第 $script:currentGroupPage 页 / 共 $script:totalGroupPages 页（总计 $($script:filteredGroups.Count) 条）"
    $script:btnGroupPrev.Enabled = ($script:currentGroupPage -gt 1)
    $script:btnGroupNext.Enabled = ($script:currentGroupPage -lt $script:totalGroupPages)
    $script:txtGroupJumpPage.Text = $script:currentGroupPage.ToString()
    $script:groupPaginationPanel.Visible = $true
}

# 保留原有的分页计算函数
function Get-TotalPages {
    param(
        [int]$totalCount,
        [int]$pageSize
    )
    if ($totalCount -eq 0) { return 0 }
    return [math]::Ceiling($totalCount / $pageSize)
}
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

    if ($script:defaultShowAll) {
        # 全显模式（不变）
        $script:userDataGridView.DataSource = $null
        $script:userDataGridView.DataSource = $script:filteredUsers
        $script:totalUserPages = 1
        $script:currentUserPage = 1
    } else {
        # 分页模式：使用动态 pageSize
        $startIndex = ($script:currentUserPage - 1) * $script:dynamicUserPageSize  # 替换为动态变量
        $endIndex = [math]::Min($startIndex + $script:dynamicUserPageSize, $script:filteredUsers.Count)  # 替换为动态变量
        $currentPageData = New-Object System.Collections.ArrayList
        for ($i = $startIndex; $i -lt $endIndex; $i++) {
            if ($null -ne $script:filteredUsers[$i]) {
                $null = $currentPageData.Add($script:filteredUsers[$i])
            }
        }
        $script:userDataGridView.DataSource = $null
        $script:userDataGridView.DataSource = $currentPageData
        # 重新计算总页数（使用动态 pageSize）
        $script:totalUserPages = Get-TotalPages -totalCount $script:filteredUsers.Count -pageSize $script:dynamicUserPageSize
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

    if ($script:groupDefaultShowAll) {
        # 全显模式（不变）
        $script:groupDataGridView.DataSource = $null
        $script:groupDataGridView.DataSource = $script:filteredGroups
        $script:totalGroupPages = 1
        $script:currentGroupPage = 1
    } else {
        # 分页模式：使用动态 pageSize（关键修改）
        $startIndex = ($script:currentGroupPage - 1) * $script:dynamicGroupPageSize  # 替换为动态变量
        $endIndex = [math]::Min($startIndex + $script:dynamicGroupPageSize, $script:filteredGroups.Count)  # 替换为动态变量
        $currentPageData = New-Object System.Collections.ArrayList
        for ($i = $startIndex; $i -lt $endIndex; $i++) {
            if ($null -ne $script:filteredGroups[$i]) {
                $null = $currentPageData.Add($script:filteredGroups[$i])
            }
        }
        $script:groupDataGridView.DataSource = $null
        $script:groupDataGridView.DataSource = $currentPageData
        # 重新计算总页数（使用动态 pageSize）
        $script:totalGroupPages = Get-TotalPages -totalCount $script:filteredGroups.Count -pageSize $script:dynamicGroupPageSize
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
    # 1. 处理总数据量为 0 的情况（直接返回 0）
    if ($totalCount -eq 0) {
        return 0
    }
    # 2. 关键修复：确保 pageSize 至少为 1（避免除以零）
    if ($pageSize -le 0) {
        $pageSize = 1  # 强制设置为 1，确保除法有效
        #Write-Host "警告：pageSize 为 0，已强制设为 1"  # 调试用
    }
    # 3. 安全计算总页数
    return [math]::Ceiling($totalCount / $pageSize)
}

# 通用函数：计算 DataGridView 可视行数（用于动态分页）
function Get-VisibleRowCount {
    param(
        [System.Windows.Forms.DataGridView]$dgv
    )
    if (-not $dgv -or $dgv.ColumnHeadersHeight -le 0 -or $dgv.RowTemplate.Height -le 0) {
        #Write-Host "DGV未就绪，返回1行"
        return 1
    }

    $dgv.PerformLayout()

    # 边框高度计算（不变）
    $borderHeight = 0
    switch ($dgv.BorderStyle) {
        [System.Windows.Forms.BorderStyle]::Fixed3D { $borderHeight = 4 }
        [System.Windows.Forms.BorderStyle]::FixedSingle { $borderHeight = 2 }
        default { $borderHeight = 0 }
    }

    # 内边距高度（不变）
    $paddingHeight = $dgv.Padding.Top + $dgv.Padding.Bottom

    # 安全冗余（不变）
    $safetyMargin = 8

    # 【修复可用高度异常】确保 ClientSize.Height 有效
    if ($dgv.ClientSize.Height -le 0) {
        #Write-Host "DGV高度无效，返回1行"
        return 1
    }

    $availableHeight = $dgv.ClientSize.Height - $dgv.ColumnHeadersHeight - $borderHeight - $paddingHeight - $safetyMargin
    # 确保可用高度至少为1行的高度（避免负数）
    $minRequiredHeight = $dgv.RowTemplate.Height  # 至少能容纳1行
    $availableHeight = [math]::Max($availableHeight, $minRequiredHeight)

    $visibleRows = [math]::Max(1, [math]::Floor($availableHeight / $dgv.RowTemplate.Height))
    # 限制最大行数为数据总量（不变）
    if ($script:filteredUsers.Count -gt 0 -and $dgv.Name -eq "userDataGridView") {
        $visibleRows = [math]::Min($visibleRows, $script:filteredUsers.Count)
    }
    if ($script:filteredGroups.Count -gt 0 -and $dgv.Name -eq "groupDataGridView") {
        $visibleRows = [math]::Min($visibleRows, $script:filteredGroups.Count)
    }
	
	$visibleRows = [math]::Max(1, $visibleRows)
    #Write-Host "计算出可视行数: $($visibleRows)"
    return $visibleRows
}


# 更新用户列表动态分页大小（添加 null 检查）
function Update-DynamicUserPageSize {
    # 1. 先检查所有面板是否已初始化
    if (-not $script:mainPanel -or -not $script:userManagementPanel -or -not $script:userListPanel) {
        #Write-Host "面板尚未初始化，跳过用户分页更新"
        return
    }

    # 2. 再执行布局刷新（确保面板已存在）
    $script:mainPanel.PerformLayout()
    $script:userManagementPanel.PerformLayout()
    $script:userListPanel.PerformLayout()
    $dgv = $script:userDataGridView
    if (-not $dgv) { return }  # 检查 DGV 是否存在
    $dgv.PerformLayout()

    $newPageSize = Get-VisibleRowCount -dgv $dgv
    if ($newPageSize -eq $script:dynamicUserPageSize) {
        Check-ScrollBar -dgv $dgv -dataCount $script:filteredUsers.Count
        return
    }
    $script:dynamicUserPageSize = $newPageSize

    $script:totalUserPages = Get-TotalPages -totalCount $script:filteredUsers.Count -pageSize $script:dynamicUserPageSize
    $script:currentUserPage = [math]::Min($script:currentUserPage, $script:totalUserPages)
    Show-UserPage

    Check-ScrollBar -dgv $dgv -dataCount $script:filteredUsers.Count
}

# 更新组列表动态分页大小（同样添加 null 检查）
function Update-DynamicGroupPageSize {
    # 1. 面板 null 检查
    if (-not $script:mainPanel -or -not $script:groupManagementPanel -or -not $script:groupListPanel) {
        #Write-Host "面板尚未初始化，跳过组分页更新"
        return
    }

    # 2. 布局刷新
    $script:mainPanel.PerformLayout()
    $script:groupManagementPanel.PerformLayout()
    $script:groupListPanel.PerformLayout()
    $dgv = $script:groupDataGridView
    if (-not $dgv) { return }  # 检查 DGV 是否存在
    $dgv.PerformLayout()

    $newPageSize = Get-VisibleRowCount -dgv $dgv
    if ($newPageSize -eq $script:dynamicGroupPageSize) {
        Check-ScrollBar -dgv $dgv -dataCount $script:filteredGroups.Count
        return
    }
    $script:dynamicGroupPageSize = $newPageSize

    $script:totalGroupPages = Get-TotalPages -totalCount $script:filteredGroups.Count -pageSize $script:dynamicGroupPageSize
    $script:currentGroupPage = [math]::Min($script:currentGroupPage, $script:totalGroupPages)
    Show-GroupPage

    Check-ScrollBar -dgv $dgv -dataCount $script:filteredGroups.Count
}


# 滚动条强制校验函数（根据实际数据高度判断）
function Check-ScrollBar {
    param(
        [System.Windows.Forms.DataGridView]$dgv,
        [int]$dataCount
    )
    if ($dataCount -eq 0) {
        $dgv.ScrollBars = [System.Windows.Forms.ScrollBars]::None
        return
    }

	$dgv.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
}
<# 
ͨ�ø������� 
#>

# ����û������������
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

# ���������������
function ClearGroupInputFields {
    $script:textGroupName.Text = ""
    $script:textGroupSamAccount.Text = ""
    $script:textGroupDescription.Text = ""
    $script:originalGroupSamAccount = $null
	$script:groupDataGridView.ClearSelection()
}

# ɸѡ�û��б�
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


# ɸѡ���б�
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

# ��ҳ��ʾ����
function Show-UserPage {
    # ��������ݣ�ԭ�߼����䣩
    if ($script:filteredUsers.Count -eq 0) {
        $script:userDataGridView.DataSource = $null
        $script:lblUserPageInfo.Text = "�� 0 ҳ / �� 0 ҳ���ܼ� 0 ����"
        $script:btnUserPrev.Enabled = $false
        $script:btnUserNext.Enabled = $false
        $script:txtUserJumpPage.Text = ""
        $script:userPaginationPanel.Visible = $false
        return
    }

    if ($script:defaultShowAll) {
        # ȫ��ģʽ�����䣩
        $script:userDataGridView.DataSource = $null
        $script:userDataGridView.DataSource = $script:filteredUsers
        $script:totalUserPages = 1
        $script:currentUserPage = 1
    } else {
        # ��ҳģʽ��ʹ�ö�̬ pageSize
        $startIndex = ($script:currentUserPage - 1) * $script:dynamicUserPageSize  # �滻Ϊ��̬����
        $endIndex = [math]::Min($startIndex + $script:dynamicUserPageSize, $script:filteredUsers.Count)  # �滻Ϊ��̬����
        $currentPageData = New-Object System.Collections.ArrayList
        for ($i = $startIndex; $i -lt $endIndex; $i++) {
            if ($null -ne $script:filteredUsers[$i]) {
                $null = $currentPageData.Add($script:filteredUsers[$i])
            }
        }
        $script:userDataGridView.DataSource = $null
        $script:userDataGridView.DataSource = $currentPageData
        # ���¼�����ҳ����ʹ�ö�̬ pageSize��
        $script:totalUserPages = Get-TotalPages -totalCount $script:filteredUsers.Count -pageSize $script:dynamicUserPageSize
    }

    # ---------------------- ���·�ҳ�ؼ�״̬ ----------------------
    $script:lblUserPageInfo.Text = "�� $script:currentUserPage ҳ / �� $script:totalUserPages ҳ���ܼ� $($script:filteredUsers.Count) ����"
    $script:btnUserPrev.Enabled = ($script:currentUserPage -gt 1)  # ֻ��ҳ��>1ʱ����һҳ
    $script:btnUserNext.Enabled = ($script:currentUserPage -lt $script:totalUserPages)  # ҳ��<��ҳ��ʱ����һҳ
    $script:txtUserJumpPage.Text = $script:currentUserPage.ToString()  # ͬ��ҳ�뵽��ת��
    $script:userPaginationPanel.Visible = $true  # ʼ����ʾ��ҳ�ؼ�
}


function Show-GroupPage {
    # ��������ݣ�ԭ�߼����䣩
    if ($script:filteredGroups.Count -eq 0) {
        $script:groupDataGridView.DataSource = $null
        $script:lblGroupPageInfo.Text = "�� 0 ҳ / �� 0 ҳ���ܼ� 0 ����"
        $script:btnGroupPrev.Enabled = $false
        $script:btnGroupNext.Enabled = $false
        $script:txtGroupJumpPage.Text = ""
        $script:groupPaginationPanel.Visible = $false
        return
    }

    if ($script:groupDefaultShowAll) {
        # ȫ��ģʽ�����䣩
        $script:groupDataGridView.DataSource = $null
        $script:groupDataGridView.DataSource = $script:filteredGroups
        $script:totalGroupPages = 1
        $script:currentGroupPage = 1
    } else {
        # ��ҳģʽ��ʹ�ö�̬ pageSize���ؼ��޸ģ�
        $startIndex = ($script:currentGroupPage - 1) * $script:dynamicGroupPageSize  # �滻Ϊ��̬����
        $endIndex = [math]::Min($startIndex + $script:dynamicGroupPageSize, $script:filteredGroups.Count)  # �滻Ϊ��̬����
        $currentPageData = New-Object System.Collections.ArrayList
        for ($i = $startIndex; $i -lt $endIndex; $i++) {
            if ($null -ne $script:filteredGroups[$i]) {
                $null = $currentPageData.Add($script:filteredGroups[$i])
            }
        }
        $script:groupDataGridView.DataSource = $null
        $script:groupDataGridView.DataSource = $currentPageData
        # ���¼�����ҳ����ʹ�ö�̬ pageSize��
        $script:totalGroupPages = Get-TotalPages -totalCount $script:filteredGroups.Count -pageSize $script:dynamicGroupPageSize
    }

    # ---------------------- ���·�ҳ�ؼ�״̬ ----------------------
    $script:lblGroupPageInfo.Text = "�� $script:currentGroupPage ҳ / �� $script:totalGroupPages ҳ���ܼ� $($script:filteredGroups.Count) ����"
    $script:btnGroupPrev.Enabled = ($script:currentGroupPage -gt 1)
    $script:btnGroupNext.Enabled = ($script:currentGroupPage -lt $script:totalGroupPages)
    $script:txtGroupJumpPage.Text = $script:currentGroupPage.ToString()
    $script:groupPaginationPanel.Visible = $true
}

# ����ԭ�еķ�ҳ���㺯��
function Get-TotalPages {
    param(
        [int]$totalCount,
        [int]$pageSize
    )
    # 1. ������������Ϊ 0 �������ֱ�ӷ��� 0��
    if ($totalCount -eq 0) {
        return 0
    }
    # 2. �ؼ��޸���ȷ�� pageSize ����Ϊ 1����������㣩
    if ($pageSize -le 0) {
        $pageSize = 1  # ǿ������Ϊ 1��ȷ��������Ч
        #Write-Host "���棺pageSize Ϊ 0����ǿ����Ϊ 1"  # ������
    }
    # 3. ��ȫ������ҳ��
    return [math]::Ceiling($totalCount / $pageSize)
}

# ͨ�ú��������� DataGridView �������������ڶ�̬��ҳ��
function Get-VisibleRowCount {
    param(
        [System.Windows.Forms.DataGridView]$dgv
    )
    if (-not $dgv -or $dgv.ColumnHeadersHeight -le 0 -or $dgv.RowTemplate.Height -le 0) {
        #Write-Host "DGVδ����������1��"
        return 1
    }

    $dgv.PerformLayout()

    # �߿�߶ȼ��㣨���䣩
    $borderHeight = 0
    switch ($dgv.BorderStyle) {
        [System.Windows.Forms.BorderStyle]::Fixed3D { $borderHeight = 4 }
        [System.Windows.Forms.BorderStyle]::FixedSingle { $borderHeight = 2 }
        default { $borderHeight = 0 }
    }

    # �ڱ߾�߶ȣ����䣩
    $paddingHeight = $dgv.Padding.Top + $dgv.Padding.Bottom

    # ��ȫ���ࣨ���䣩
    $safetyMargin = 8

    # ���޸����ø߶��쳣��ȷ�� ClientSize.Height ��Ч
    if ($dgv.ClientSize.Height -le 0) {
        #Write-Host "DGV�߶���Ч������1��"
        return 1
    }

    $availableHeight = $dgv.ClientSize.Height - $dgv.ColumnHeadersHeight - $borderHeight - $paddingHeight - $safetyMargin
    # ȷ�����ø߶�����Ϊ1�еĸ߶ȣ����⸺����
    $minRequiredHeight = $dgv.RowTemplate.Height  # ����������1��
    $availableHeight = [math]::Max($availableHeight, $minRequiredHeight)

    $visibleRows = [math]::Max(1, [math]::Floor($availableHeight / $dgv.RowTemplate.Height))
    # �����������Ϊ�������������䣩
    if ($script:filteredUsers.Count -gt 0 -and $dgv.Name -eq "userDataGridView") {
        $visibleRows = [math]::Min($visibleRows, $script:filteredUsers.Count)
    }
    if ($script:filteredGroups.Count -gt 0 -and $dgv.Name -eq "groupDataGridView") {
        $visibleRows = [math]::Min($visibleRows, $script:filteredGroups.Count)
    }
	
	$visibleRows = [math]::Max(1, $visibleRows)
    #Write-Host "�������������: $($visibleRows)"
    return $visibleRows
}


# �����û��б�̬��ҳ��С����� null ��飩
function Update-DynamicUserPageSize {
    # 1. �ȼ����������Ƿ��ѳ�ʼ��
    if (-not $script:mainPanel -or -not $script:userManagementPanel -or -not $script:userListPanel) {
        #Write-Host "�����δ��ʼ���������û���ҳ����"
        return
    }

    # 2. ��ִ�в���ˢ�£�ȷ������Ѵ��ڣ�
    $script:mainPanel.PerformLayout()
    $script:userManagementPanel.PerformLayout()
    $script:userListPanel.PerformLayout()
    $dgv = $script:userDataGridView
    if (-not $dgv) { return }  # ��� DGV �Ƿ����
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

# �������б�̬��ҳ��С��ͬ����� null ��飩
function Update-DynamicGroupPageSize {
    # 1. ��� null ���
    if (-not $script:mainPanel -or -not $script:groupManagementPanel -or -not $script:groupListPanel) {
        #Write-Host "�����δ��ʼ�����������ҳ����"
        return
    }

    # 2. ����ˢ��
    $script:mainPanel.PerformLayout()
    $script:groupManagementPanel.PerformLayout()
    $script:groupListPanel.PerformLayout()
    $dgv = $script:groupDataGridView
    if (-not $dgv) { return }  # ��� DGV �Ƿ����
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


# ������ǿ��У�麯��������ʵ�����ݸ߶��жϣ�
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
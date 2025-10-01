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

    # ---------------------- �������ж��Ƿ�Ĭ��ȫ�� ----------------------
    if ($script:defaultShowAll) {
        # ȫ��ģʽ����ʾ�������ݣ���ҳ�ؼ����÷�ҳ
        $script:userDataGridView.DataSource = $null
        $script:userDataGridView.DataSource = $script:filteredUsers
        $script:totalUserPages = 1  # ȫ��ʱ��ҳ��=1
        $script:currentUserPage = 1  # ȫ��ʱҳ��=1
    }
    else {
        # ��ҳģʽ����ԭ�߼���ȡ��ǰҳ����
        $startIndex = ($script:currentUserPage - 1) * $script:pageSize
        $endIndex = [math]::Min($startIndex + $script:pageSize, $script:filteredUsers.Count)
        $currentPageData = New-Object System.Collections.ArrayList
        for ($i = $startIndex; $i -lt $endIndex; $i++) {
            if ($null -ne $script:filteredUsers[$i]) {
                $null = $currentPageData.Add($script:filteredUsers[$i])
            }
        }
        # �󶨷�ҳ����
        $script:userDataGridView.DataSource = $null
        $script:userDataGridView.DataSource = $currentPageData
        # ������ҳ������ҳģʽ����pageSize���㣩
        $script:totalUserPages = Get-TotalPages -totalCount $script:filteredUsers.Count -pageSize $script:pageSize
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

    # ---------------------- �������ж��Ƿ�Ĭ��ȫ�� ----------------------
    if ($script:groupDefaultShowAll) {
        # ȫ��ģʽ����ʾ��������
        $script:groupDataGridView.DataSource = $null
        $script:groupDataGridView.DataSource = $script:filteredGroups
        $script:totalGroupPages = 1
        $script:currentGroupPage = 1
    }
    else {
        # ��ҳģʽ����ȡ��ǰҳ����
        $startIndex = ($script:currentGroupPage - 1) * $script:pageSize
        $endIndex = [math]::Min($startIndex + $script:pageSize, $script:filteredGroups.Count)
        $currentPageData = New-Object System.Collections.ArrayList
        for ($i = $startIndex; $i -lt $endIndex; $i++) {
            if ($null -ne $script:filteredGroups[$i]) {
                $null = $currentPageData.Add($script:filteredGroups[$i])
            }
        }
        # �󶨷�ҳ����
        $script:groupDataGridView.DataSource = $null
        $script:groupDataGridView.DataSource = $currentPageData
        # ������ҳ��
        $script:totalGroupPages = Get-TotalPages -totalCount $script:filteredGroups.Count -pageSize $script:pageSize
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
    if ($totalCount -eq 0) { return 0 }
    return [math]::Ceiling($totalCount / $pageSize)
}
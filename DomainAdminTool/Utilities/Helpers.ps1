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



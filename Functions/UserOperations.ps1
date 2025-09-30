<# 
�û���غ��Ĳ��� 
#>

function LoadUserList {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    try {
        $script:connectionStatus = "���ڼ��� OU: $($script:currentOU) �µ��û�..."
        UpdateStatusBar
        $script:mainForm.Refresh()
        
        $script:allUsers.Clear()
        $script:filteredUsers.Clear()

        # Զ�̼����û����߼����䣩
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

            # �����û����ݣ��߼����䣩
            $users | ForEach-Object {
                $groupNames = $_.MemberOf | ForEach-Object { if ($_ -match 'CN=([^,]+)') { $matches[1] } }
                $groupsString = if ($groupNames) { $groupNames -join ', ' } else { '��' }
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

        # ������ݣ��߼����䣩
        $remoteUsers | ForEach-Object {
            $null = $script:allUsers.Add($_)
            $null = $script:filteredUsers.Add($_)
        }
        
        # ---------------------- �ؼ���Ĭ��ȫ������ ----------------------
        $script:defaultShowAll = $true  # �л�OU��ǿ��Ĭ��ȫ��
        $script:currentUserPage = 1  # ���õ�ǰҳ��Ϊ1
        # ȫ��ʱ��ҳ��=1��������������ΪpageSize��
        $script:totalUserPages = Get-TotalPages -totalCount $script:filteredUsers.Count -pageSize $script:filteredUsers.Count  
        
        # 1. ��ȫ�����ݵ�DataGridView
        $script:userDataGridView.DataSource = $null
        $script:userDataGridView.DataSource = $script:filteredUsers
        
        # 2. ͬ����ҳ�ؼ�״̬�����÷�ҳ��ť����ʾ��ȷҳ�룩
        $script:lblUserPageInfo.Text = "�� $script:currentUserPage ҳ / �� $script:totalUserPages ҳ���ܼ� $($script:filteredUsers.Count) ����"
        $script:btnUserPrev.Enabled = $false  # ȫ��ʱ�޷���һҳ
        $script:btnUserNext.Enabled = $false  # ȫ��ʱ�޷���һҳ
        $script:txtUserJumpPage.Text = "1"  # ��ת��Ĭ����ʾ1
        $script:userPaginationPanel.Visible = $true  # ǿ����ʾ��ҳ�ؼ�
        # ----------------------------------------------------------------
        
        # ����״̬���߼����䣩
        $script:userCountStatus = $script:allUsers.Count
        $script:connectionStatus = "�Ѽ��� OU: $($script:currentOU) �µ� $($script:userCountStatus) ���û�"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $script:connectionStatus = "�����û�ʧ��: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
    }
}



function ToggleUserEnabled {
    param($rowIndex)
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $script:userDataGridView.CancelEdit()
        return
    }
    
    $selectedRow = $script:userDataGridView.Rows[$rowIndex]
    $enabledCell = $selectedRow.Cells["Enabled"]
    
    # ��ȡ��ǰ״̬
    $currentState = $false
    if ($enabledCell.Value -ne $null) { $currentState = [bool]$enabledCell.Value }
    $newState = -not $currentState
    $action = if ($newState) { "����" } else { "����" }
    
    # ��ȡ�˺�
    $username = $null
    if ($selectedRow.DataBoundItem -ne $null) {
        $username = $selectedRow.DataBoundItem.SamAccountName.ToString().Trim()
    }
    if ([string]::IsNullOrEmpty($username)) {
        $username = $selectedRow.Cells["SamAccountName"].Value.ToString().Trim()
    }
    if ([string]::IsNullOrEmpty($username)) {
        [System.Windows.Forms.MessageBox]::Show("δ�ҵ��˺���Ϣ", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $script:userDataGridView.CancelEdit()
        return
    }
    
    # ȷ�ϲ���
    if ([System.Windows.Forms.MessageBox]::Show("ȷ��Ҫ$($action)�˺� [$username] ��", "ȷ��", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne 'Yes') {
        $script:userDataGridView.CancelEdit()
        return
    }
    
    try {
        $script:connectionStatus = "����$($action)�˺�..."
        UpdateStatusBar
        $script:mainForm.Refresh()
        
        # Զ��ִ������/����
        $remoteResult = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($username, $newState)
            Import-Module ActiveDirectory -ErrorAction Stop
            $user = Get-ADUser -Filter { SamAccountName -eq $username } -ErrorAction Stop
            Set-ADUser -Identity $user.DistinguishedName -Enabled $newState -ErrorAction Stop
            $updatedUser = Get-ADUser -Identity $user.DistinguishedName -Properties Enabled -ErrorAction Stop
            return $updatedUser.Enabled
        } -ArgumentList $username, $newState -ErrorAction Stop
        
        if ($remoteResult -ne $newState) { throw "������״̬δ���" }
        [System.Windows.Forms.MessageBox]::Show("�˺� [$username] $($action)�ɹ�", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "${action}ʧ��: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        $script:userDataGridView.CancelEdit()
        [System.Windows.Forms.MessageBox]::Show("�˺� [$username] $($action)ʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}



function CreateNewUser {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    # ��ȡ���벢ȷ������������Ϊ��
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

    # ��ǿУ�飬ȷ�����б�Ҫ��������ֵ
    if ([string]::IsNullOrWhiteSpace($cnName)) {
        [System.Windows.Forms.MessageBox]::Show("��������Ϊ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ([string]::IsNullOrWhiteSpace($username)) {
        [System.Windows.Forms.MessageBox]::Show("�˺Ų���Ϊ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ([string]::IsNullOrWhiteSpace($password)) {
        [System.Windows.Forms.MessageBox]::Show("���벻��Ϊ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($password -ne $confirm) {
        [System.Windows.Forms.MessageBox]::Show("�������벻һ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($password.Length -lt 8 -or $password -notmatch '[A-Z]' -or $password -notmatch '[a-z]' -or $password -notmatch '[0-9]' -or $password -notmatch '[^a-zA-Z0-9]') {
        [System.Windows.Forms.MessageBox]::Show("�������8λ��������Сд��ĸ�����ֺ������ַ�����@#��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ([string]::IsNullOrWhiteSpace($currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("��ѡ���û���֯��λ(OU)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ([string]::IsNullOrWhiteSpace($domainDNSRoot)) {
        [System.Windows.Forms.MessageBox]::Show("�޷���ȡ����Ϣ������������", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    try {
        $script:connectionStatus = "����Զ�̴����˺�[$username]..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # ȷ��Զ�̻Ự����
        if (-not $script:remoteSession -or $script:remoteSession.State -ne 'Opened') {
            throw "Զ�̻Ựδ���ӣ����������ӵ����"
        }

        # Զ�̴����û�
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

                # ȷ��ActiveDirectoryģ���Ѽ���
                if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
                    throw "ActiveDirectoryģ��δ��װ"
                }
                Import-Module ActiveDirectory -ErrorAction Stop

                # ���帴���б�
				$doubleSurnames = @(
					"ŷ��", "̫ʷ", "��ľ", "�Ϲ�", "˾��", "����", "����", "�Ϲ�",
					"��ٹ", "����", "�ĺ�", "���", "ξ��", "����", "����", "�̨",
					"�ʸ�", "����", "���", "��ұ", "̫��", "����", "����", "Ľ��",
					"����", "����", "˾ͽ", "����", "˾��", "����", "˾��", "�붽",
					"�ӳ�", "���", "��ľ", "����", "����", "���", "����", "����",
					"����", "�ذ�", "�й�", "�׸�", "����", "�θ�", "����", "����",
					"����", "����", "����", "΢��", "����", "����", "����", "����",
					"�Ը�", "����", "����", "����", "��ī", "����", "��ʦ", "����",
					"����", "����", "�ٳ�", "��ͻ", "����", "����", "Ľ��", "ξ��",
					"��Ƶ", "������", "����", "����", "����", "߳��", "����", "ͺ��",
					"���", "�ǵ�", "�ڹ���", "����", "�й�", "�Ѳ�", "Ů����", "أ��",
					"˹��", "�ﲮ", "�麣", "����", "����", "����", "ұ��", "����"
				)

                # ��ȡ�պ�����֧�ָ���
                if ($cnName.Length -ge 2 -and $doubleSurnames -contains $cnName.Substring(0, 2)) {
                    $surname = $cnName.Substring(0, 2)
                    $givenName = if ($cnName.Length -gt 2) { $cnName.Substring(2) } else { "" }
                } else {
                    $surname = $cnName.Substring(0, 1)
                    $givenName = if ($cnName.Length -gt 1) { $cnName.Substring(1) } else { $cnName }
                }

                # ������ȫ����
                $securePwd = ConvertTo-SecureString $password -AsPlainText -Force -ErrorAction Stop

                # ׼���û�����
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
                # �����˺Ź���ʱ��
                if (-not $neverExpire -and $expiryDate -gt (Get-Date)) {
                    $userParams.AccountExpirationDate = $expiryDate
                }
#>
                if (-not $neverExpire) {
                    $userParams.AccountExpirationDate = $expiryDate
                }


                # �����û� - ���û����Զ�����Domain Users��
                New-ADUser @userParams

                # ��ȡ�������û���Ϣ
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

        [System.Windows.Forms.MessageBox]::Show("�˺Ŵ����ɹ���`n�˺ţ�$username`n������$cnName`n����·����$($remoteResult.DN)", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
        ClearInputFields  # ����Helpers.ps1
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "����ʧ��: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("�˺�[$username]����ʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}




function ModifyUserAccount {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("��ѡ����Ҫ�޸ĵ��û�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ��ȡѡ���˺���Ϣ
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
        [System.Windows.Forms.MessageBox]::Show("δ�ҵ�ԭ�˺���Ϣ", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ��ȡ���������Ϣ
    $newDisplayName = $script:textCnName.Text.Trim()
    $newEmail = $script:textEmail.Text.Trim()
	$newPhone = $script:textPhone.Text.Trim()
    $newDescription = $script:textDescription.Text.Trim()
    $neverExpire = $script:chkNeverExpire.Checked
    $expiryDate = $script:dateExpiry.Value.AddDays(1)

    # ��֤�Ƿ����޸�
	$oldPhone = if ($selectedRow.Cells["OfficePhone"] -ne $null) { $selectedRow.Cells["OfficePhone"].Value } else { "" }
    if ($newDisplayName -eq $oldDisplayName -and 
        $newEmail -eq $selectedRow.Cells["EmailAddress"].Value -and
        $newPhone -eq $oldPhone -and		
        $newDescription -eq $selectedRow.Cells["Description"].Value) {
        [System.Windows.Forms.MessageBox]::Show("δ��⵽�κ��޸�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    try {
        $script:connectionStatus = "����Զ���޸��˺�[$oldUsername]��Ϣ..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # Զ��ִ���޸�
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

        [System.Windows.Forms.MessageBox]::Show("�˺�[$oldUsername]��Ϣ�޸ĳɹ�", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
        $script:connectionStatus = "�����ӵ����: $($script:comboDomain.SelectedItem.Name)��Զ��ִ�У�"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "�޸�ʧ��: $errorMsg"
		Write-Error $errorMsg
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("�˺�[$oldUsername]��Ϣ�޸�ʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}
    



function ChangeUserPassword {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("��ѡ���û�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    # ��ȡѡ���˺�
    $selectedRow = $script:userDataGridView.SelectedRows[0]
    $username = $selectedRow.Cells["SamAccountName"].Value.ToString().Trim()
    
    $password = $script:textNewPassword.Text
    $confirm = $script:textConfirmPassword.Text
    
    if (-not $password) {
        [System.Windows.Forms.MessageBox]::Show("������������", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($password -ne $confirm) {
        [System.Windows.Forms.MessageBox]::Show("�������벻һ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($password.Length -lt 8 -or $password -notmatch '[A-Z]' -or $password -notmatch '[a-z]' -or $password -notmatch '[0-9]' -or $password -notmatch '[^a-zA-Z0-9]') {
        [System.Windows.Forms.MessageBox]::Show("�������8λ��������Сд��ĸ�����ֺ������ַ�", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    try {
        $script:connectionStatus = "����Զ���޸�����..."
        UpdateStatusBar
        $script:mainForm.Refresh()
        
        # Զ��ִ�������޸�
        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($username, $password)
            Import-Module ActiveDirectory -ErrorAction Stop
            $securePwd = ConvertTo-SecureString $password -AsPlainText -Force
            Set-ADAccountPassword -Identity $username -NewPassword $securePwd -Reset -ErrorAction Stop
            Set-ADUser -Identity $username -ChangePasswordAtLogon $true -ErrorAction Stop
        } -ArgumentList $username, $password -ErrorAction Stop
        
        [System.Windows.Forms.MessageBox]::Show("�����޸ĳɹ�", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
        $script:textNewPassword.Text = ""
        $script:textConfirmPassword.Text = ""
    }
    catch {
        $script:connectionStatus = "�޸�ʧ��: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("�޸�ʧ��: $($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function UnlockUserAccount {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("��ѡ���û�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    # ��ȡѡ���˺���Ϣ
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
        [System.Windows.Forms.MessageBox]::Show("δ�ҵ��˺���Ϣ��������", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Զ�̲�ѯ�˺�����״̬
    try {
        $script:connectionStatus = "���ڲ�ѯ�˺�����״̬..."
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
        $script:connectionStatus = "��ѯ����״̬ʧ��: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("��ѯ�˺�[$username]״̬ʧ�ܣ�`n$($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ��֤�Ƿ���Ҫ����
    if (-not $isLocked) {
        [System.Windows.Forms.MessageBox]::Show("�˺�[$username]δ�������������", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ȷ�Ͻ�������
    if ([System.Windows.Forms.MessageBox]::Show("ȷ�������˺� [$username] ��", "ȷ�Ͻ���", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne 'Yes') {
        return
    }

    # Զ��ִ�н���
    try {
        $script:connectionStatus = "����Զ�̽����˺�[$username]..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($targetUser)
            Import-Module ActiveDirectory -ErrorAction Stop
            $user = Get-ADUser -Identity $targetUser -ErrorAction Stop
            Unlock-ADAccount -Identity $user.DistinguishedName -ErrorAction Stop
            $updatedUser = Get-ADUser -Identity $user.DistinguishedName -Properties LockedOut -ErrorAction Stop
            if ($updatedUser.LockedOut) {
                throw "�������˺��Դ�������״̬�����������ͬ���ӳ�"
            }
        } -ArgumentList $username -ErrorAction Stop

        [System.Windows.Forms.MessageBox]::Show("�˺�[$username]�����ɹ�", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
        $script:connectionStatus = "�����ӵ����: $($script:comboDomain.SelectedItem.Name)��Զ��ִ�У�"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "����ʧ��: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("�˺�[$username]����ʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function RenameUserAccount {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("��ѡ����Ҫ���������û�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ��ȡԭ�˺���Ϣ
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
        [System.Windows.Forms.MessageBox]::Show("δ�ҵ�ԭ�˺���Ϣ", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ��ȡ���˺���Ϣ
    $newUsername = $script:textPinyin.Text.Trim()
    $newDisplayName = $script:textCnName.Text.Trim()

    if ([string]::IsNullOrEmpty($newUsername) -or [string]::IsNullOrEmpty($newDisplayName)) {
        [System.Windows.Forms.MessageBox]::Show("���������˺ź�������", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($newUsername -eq $oldUsername -and $newDisplayName -eq $oldDisplayName) {
        [System.Windows.Forms.MessageBox]::Show("����Ϣ��ԭ��Ϣһ�£������޸�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # Զ�̼�����˺��Ƿ��Ѵ���
    try {
        $script:connectionStatus = "���ڼ�����˺ſ�����..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        $exists = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($newUser)
            Import-Module ActiveDirectory -ErrorAction Stop
            $user = Get-ADUser -Filter "SamAccountName -eq '$newUser'" -ErrorAction SilentlyContinue
            return $null -ne $user
        } -ArgumentList $newUsername -ErrorAction Stop

        if ($exists) {
            [System.Windows.Forms.MessageBox]::Show("���˺�[$newUsername]�Ѵ��ڣ������", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }
    catch {
        $script:connectionStatus = "����˺�ʧ��: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("������˺ſ�����ʧ�ܣ�`n$($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ȷ������������
    $confirmMsg = "ȷ���������˺ţ�`nԭ��Ϣ��$oldDisplayName��$oldUsername��`n����Ϣ��$newDisplayName��$newUsername��"
    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, "ȷ��������", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne 'Yes') {
        return
    }

    # Զ��ִ��������
    try {
        $script:connectionStatus = "����Զ���������˺�[$oldUsername]..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($oldUser, $newUser, $newDisplayName, $domainDNSRoot)
            Import-Module ActiveDirectory -ErrorAction Stop

            $user = Get-ADUser -Identity $oldUser -Properties DistinguishedName -ErrorAction Stop
            $userDN = $user.DistinguishedName

            # �޸�SamAccountName��UPN
            Set-ADUser -Identity $userDN `
                       -SamAccountName $newUser `
                       -UserPrincipalName "$newUser@$domainDNSRoot" `
                       -ErrorAction Stop

            # �޸�DisplayName
            Set-ADUser -Identity $newUser `
                       -DisplayName $newDisplayName `
                       -GivenName $newDisplayName `
                       -ErrorAction Stop

            # ��֤�޸Ľ��
            $updatedUser = Get-ADUser -Identity $newUser -Properties DisplayName, UserPrincipalName -ErrorAction Stop
            if ($updatedUser.SamAccountName -ne $newUser -or $updatedUser.DisplayName -ne $newDisplayName) {
                throw "����������Ϣ��ƥ�䣬�޸�δ��ȫ��Ч"
            }
        } -ArgumentList $oldUsername, $newUsername, $newDisplayName, $script:domainContext.DomainInfo.DNSRoot -ErrorAction Stop

        [System.Windows.Forms.MessageBox]::Show("�˺��������ɹ���`n���˺ţ�$newUsername`n��������$newDisplayName", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadUserList
        ClearInputFields
        $script:connectionStatus = "�����ӵ����: $($script:comboDomain.SelectedItem.Name)��Զ��ִ�У�"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "������ʧ��: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("�˺�������ʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}



function DeleteUserAccount {
    # 1. ����У�飺�Ƿ��������
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    # ��ȡ����ѡ����
    $selectedRows = $script:userDataGridView.SelectedRows
    if ($selectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("��ͨ��Ctrl����ѡ��Ҫɾ�����û�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }


    # 2. ��ȡ����ѡ���е��û���Ϣ�����ӿ�ֵ��飩
    $selectedUsers = @()
    foreach ($row in $selectedRows) {
        $userInfo = [PSCustomObject]@{
            DisplayName        = ""  # Ĭ��Ϊ���ַ�������null
            SamAccountName     = ""
            IsValid            = $false
            DistinguishedName  = $null
        }

        # ��DataBoundItem��ȡ�����ȷ�ʽ��
        if ($row.DataBoundItem -ne $null) {
            $userData = $row.DataBoundItem
            
            # ��ȫ����DisplayName������null���÷�����
            if ($null -ne $userData.DisplayName) {
                $userInfo.DisplayName = $userData.DisplayName.ToString().Trim()
            }
            
            # ��ȫ����SamAccountName���˺��ǹؼ���Ϣ������У�飩
            if ($null -ne $userData.SamAccountName) {
                $userInfo.SamAccountName = $userData.SamAccountName.ToString().Trim()
            }
        }
        # �ӵ�Ԫ��ֱ����ȡ�����÷�ʽ��
        else {
            # ����DisplayName��Ԫ��
            if ($row.Cells["DisplayName"] -ne $null -and $null -ne $row.Cells["DisplayName"].Value) {
                $userInfo.DisplayName = $row.Cells["DisplayName"].Value.ToString().Trim()
            }
            
            # ����SamAccountName��Ԫ�񣨹ؼ���Ϣ��
            if ($row.Cells["SamAccountName"] -ne $null -and $null -ne $row.Cells["SamAccountName"].Value) {
                $userInfo.SamAccountName = $row.Cells["SamAccountName"].Value.ToString().Trim()
            }
        }

        # ֻ��������Ч�˺ŵ��û�
        if (-not [string]::IsNullOrEmpty($userInfo.SamAccountName)) {
            $selectedUsers += $userInfo
        }
    }

    # У����Ч�û�����
    if ($selectedUsers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("ѡ�е�����δ�ҵ���Ч�˺���Ϣ��������ѡ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }


    # 3. ������֤�˺�������е���Ч��
    $script:connectionStatus = "������֤ $($selectedUsers.Count) ���˺ŵ���Ч��..."
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
            $script:connectionStatus = "�˺���֤ʧ��: $($user.SamAccountName)"
            UpdateStatusBar
            [System.Windows.Forms.MessageBox]::Show(
                "�˺� [$($user.SamAccountName)] ��֤ʧ�ܣ�`n$errorMsg", 
                "��֤����", 
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


    # 4. ����ɾ������ȷ��
    $confirmMsg = "ȷ������ɾ������ $($selectedUsers.Count) ���˺ţ�`n`n"
    $confirmMsg += "��� | ���� | �˺�`n"
    $confirmMsg += "------------------------`n"
    for ($i = 0; $i -lt $selectedUsers.Count; $i++) {
        $user = $selectedUsers[$i]
        $confirmMsg += "$($i+1).   | $($user.DisplayName) | $($user.SamAccountName)`n"
    }
    $confirmMsg += "`n�˲������ɻָ���ɾ�����޷��ָ����ݣ�"

    if ([System.Windows.Forms.MessageBox]::Show(
        $confirmMsg, 
        "����ɾ������", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) -ne 'Yes') {
        return
    }


    # 5. ִ������ɾ��
    $script:connectionStatus = "����ִ������ɾ��..."
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
                # ����У���˺Ŵ�����
                $adUser = Get-ADUser -Filter "DistinguishedName -eq '$targetDN'" -ErrorAction Stop
                Remove-ADUser -Identity $targetDN -Confirm:$false -ErrorAction Stop
                # ��֤ɾ�����
                $remaining = Get-ADUser -Filter "DistinguishedName -eq '$targetDN'" -ErrorAction SilentlyContinue
                if ($remaining) { throw "ɾ�������ܲ�ѯ���˺ţ�������AD���棩" }
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


    # 6. ��ʾɾ�����
    $resultMsg = "ɾ����ɣ�`n`n"
    $resultMsg += "�ɹ�ɾ����$($deleteResults.SuccessCount) ���˺�`n"
    $resultMsg += "ɾ��ʧ�ܣ�$($deleteResults.FailedCount) ���˺�`n"

    if ($deleteResults.FailedCount -gt 0) {
        $resultMsg += "`nʧ�����飺`n"
        foreach ($failed in $deleteResults.FailedUsers) {
            $resultMsg += "- $($failed.Account)��$($failed.Name)����$($failed.ErrorMsg)`n"
        }
        [System.Windows.Forms.MessageBox]::Show($resultMsg, "�������", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $script:connectionStatus = "����ɾ����ɣ�����ʧ�ܣ�"
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::Orange
    }
    else {
        [System.Windows.Forms.MessageBox]::Show($resultMsg, "�����ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $script:connectionStatus = "����ɾ���ɹ�"
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }


    # 7. ˢ�½���
    $script:userDataGridView.ClearSelection()
    LoadUserList  # ˢ���û��б�
    UpdateStatusBar
}







<# 
����غ��Ĳ��� 
#>

function LoadGroupList {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    try {
        $script:connectionStatus = "���ڼ��� OU: $($script:currentOU) �µ���..."
        UpdateStatusBar
        $script:mainForm.Refresh()
        
        $script:allGroups.Clear()
        $script:filteredGroups.Clear()
        
        # Զ�̼����飨�߼����䣩
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
        
        # ������ݣ��߼����䣩
        $remoteGroups | ForEach-Object {
            $null = $script:allGroups.Add($_)
            $null = $script:filteredGroups.Add($_)
        }
        
        # ---------------------- �ؼ���Ĭ��ȫ������ ----------------------
        $script:groupDefaultShowAll = $true  # �л�OU��ǿ��Ĭ��ȫ��
        $script:currentGroupPage = 1  # ���õ�ǰҳ��Ϊ1
        # ȫ��ʱ��ҳ��=1��������������ΪpageSize��
        $script:totalGroupPages = Get-TotalPages -totalCount $script:filteredGroups.Count -pageSize $script:filteredGroups.Count  
        
        # 1. ��ȫ�����ݵ�DataGridView
        $script:groupDataGridView.DataSource = $null
        $script:groupDataGridView.DataSource = $script:filteredGroups
        
        # 2. ͬ����ҳ�ؼ�״̬
        $script:lblGroupPageInfo.Text = "�� $script:currentGroupPage ҳ / �� $script:totalGroupPages ҳ���ܼ� $($script:filteredGroups.Count) ����"
        $script:btnGroupPrev.Enabled = $false
        $script:btnGroupNext.Enabled = $false
        $script:txtGroupJumpPage.Text = "1"
        $script:groupPaginationPanel.Visible = $true
        # ----------------------------------------------------------------
        
		# �״μ��غ��ʼ����̬��ҳ��С
		Update-DynamicGroupPageSize
		
        # ����״̬���߼����䣩
        $script:groupCountStatus = $script:allGroups.Count
        $script:connectionStatus = "�Ѽ��� OU: $($script:currentOU) �µ� $($script:groupCountStatus) ����"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "������ʧ��: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("���б����ʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


function CreateNewGroup {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ��ȡ������Ϣ
    $groupName = $script:textGroupName.Text.Trim()
    $groupSam = $script:textGroupSamAccount.Text.Trim()
    $groupDesc = $script:textGroupDescription.Text.Trim()

    if ([string]::IsNullOrEmpty($groupName) -or [string]::IsNullOrEmpty($groupSam)) {
        [System.Windows.Forms.MessageBox]::Show("�����ƺ����˺�Ϊ������", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Զ�̼�����Ƿ��Ѵ���
    try {
        $script:connectionStatus = "���ڼ���������..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        $exists = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($groupSamAccount, $NameOU)
            Import-Module ActiveDirectory -ErrorAction Stop
            $group = Get-ADGroup -Filter "SamAccountName -eq '$groupSamAccount'" -ErrorAction SilentlyContinue
            return $null -ne $group
        } -ArgumentList $groupSam , $script:currentOU -ErrorAction Stop

        if ($exists) {
            [System.Windows.Forms.MessageBox]::Show("���˺�[$groupSam]�Ѵ��ڣ������", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }
    catch {
        $script:connectionStatus = "�����ʧ��: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("����������ʧ�ܣ�`n$($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Զ�̴�����
    try {
        $script:connectionStatus = "����Զ�̴�����[$groupName]..."
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
                throw "�鴴���ɹ��������Բ�ƥ��"
            }
        } -ArgumentList $groupName, $groupSam, $groupDesc, $script:domainContext.DomainInfo.DefaultPartition -ErrorAction Stop

        [System.Windows.Forms.MessageBox]::Show("��[$groupName]�����ɹ�", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        LoadGroupList
        ClearGroupInputFields  # ����Helpers.ps1
        $script:connectionStatus = "�����ӵ����: $($script:comboDomain.SelectedItem.Name)��Զ��ִ�У�"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "������ʧ��: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("�鴴��ʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


function AddUserToGroup {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ֧�ֶ�ѡ�û�������Ƿ�ѡ������1���û������ݵ�/��ѡ��
    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("����ѡ����Ҫ��������û���֧��Ctrl��ѡ��", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ����Ƿ�ѡ��Ŀ���飨�����飩
    if ($script:groupDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("����ѡ��Ŀ���飨��֧�ֵ����飩", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $selectedGroup = $script:groupDataGridView.SelectedRows[0].DataBoundItem
    if (-not $selectedGroup) {
        [System.Windows.Forms.MessageBox]::Show("ѡ���������쳣��������ѡ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ��ȡ����ѡ���û��ĺ�����Ϣ��SamAccountName + ��ʾ����
    $selectedUsers = @()
    foreach ($userRow in $script:userDataGridView.SelectedRows) {
        # ��ȡ�û�SamAccountName��AD����Ψһ��ʶ������ǿգ�
        if ($userRow.Cells["SamAccountName"].Value -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("ѡ���û��������쳣���˺�Ϊ�գ���������ѡ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $username = $userRow.Cells["SamAccountName"].Value.ToString().Trim()

        # ��ȡ�û���ʾ����������ʾ��Ϊ�������˺Ŵ��棩
        $userDisplay = if ($userRow.Cells["DisplayName"].Value -ne $null) {
            $userRow.Cells["DisplayName"].Value.ToString().Trim()
        } else {
            $username
        }

        # ����������û��б�
        $selectedUsers += [PSCustomObject]@{
            SamAccountName = $username
            DisplayName    = $userDisplay
        }
    }

    # ��ȡĿ������Ϣ
    if ($selectedGroup.SamAccountName -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("ѡ������˺���ϢΪ�գ�������ѡ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $groupSam = $selectedGroup.SamAccountName.ToString().Trim()
    $groupName = if ($selectedGroup.Name -ne $null) {
        $selectedGroup.Name.ToString().Trim()
    } else {
        $groupSam
    }

    # ��������û��Ƿ���������
    $existingUsers = @()  # �������е��û�
    $validUsers = @()     # ����ӵ���Ч�û�
    try {
        $script:connectionStatus = "���ڼ�� $($selectedUsers.Count) ���û������Ա��ϵ..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # Զ��������ȡ�������г�Ա������AD���ô������������ܣ�
        $groupMembers = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($targetGroup)
            Import-Module ActiveDirectory -ErrorAction Stop
            $members = Get-ADGroupMember -Identity $targetGroup -Recursive -ErrorAction Stop
            return $members.SamAccountName  # ������Sam�˺ţ��������ݴ���
        } -ArgumentList $groupSam -ErrorAction Stop

        # �Ա�ɸѡ���Ѵ��ڵ��û� vs ����ӵ��û�
        foreach ($user in $selectedUsers) {
            if ($groupMembers -contains $user.SamAccountName) {
                $existingUsers += $user
            } else {
                $validUsers += $user
            }
        }

        # ��ʾ�Ѵ��ڵ��û������жϲ���������֪��
        if ($existingUsers.Count -gt 0) {
            $existingNames = @()
            foreach ($user in $existingUsers) {
                $existingNames += "$($user.DisplayName)��$($user.SamAccountName)��"
            }
            [System.Windows.Forms.MessageBox]::Show("�����û�������[$groupName]�У������ظ���ӣ�`n`n$($existingNames -join "`n")", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }

        # �������û����Ѵ��ڣ�ֱ���˳�
        if ($validUsers.Count -eq 0) {
            $script:connectionStatus = "�����ӵ����: $($script:comboDomain.SelectedItem.Name)��Զ��ִ�У�"
            UpdateStatusBar
            return
        }
    }
    catch {
        $script:connectionStatus = "����Ա��ϵʧ��: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("����û����ϵʧ�ܣ�`n$($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ����ȷ����Ӳ��������ֵ�/���û���ʾ�İ���
    $validUserNames = @()
    foreach ($user in $validUsers) {
        $validUserNames += "$($user.DisplayName)��$($user.SamAccountName)��"
    }
    
    $confirmTitle = if ($validUsers.Count -eq 1) { "ȷ������û�����" } else { "ȷ����������û�����" }
    $confirmMsg = if ($validUsers.Count -eq 1) {
        "ȷ�����û�`n`n$($validUserNames -join "`n")`n`n������[$groupName��$groupSam��]��"
    } else {
        "��ѡ�� $($validUsers.Count) ���û���ȷ������������[$groupName��$groupSam��]��`n`n������û���`n$($validUserNames -join "`n")"
    }

    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, $confirmTitle, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne 'Yes') {
        return
    }

    # ����Զ������û����飨һ�δ��������û�Sam�˺ţ�����AD���ã�
    try {
        $script:connectionStatus = "����Զ����� $($validUsers.Count) ���û�����[$groupName]..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # ��ȡ������û���Sam�˺��б�AD����������˸�ʽ��
        $validUserSams = $validUsers.SamAccountName

        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($users, $group)
            Import-Module ActiveDirectory -ErrorAction Stop
            # ������ӣ�Add-ADGroupMember֧�ֶ��Ա������
            Add-ADGroupMember -Identity $group -Members $users -ErrorAction Stop
            
            # ��֤��ȷ�������û����Ѽ���
            $updatedMembers = Get-ADGroupMember -Identity $group -Recursive -ErrorAction Stop
            $missingUsers = $users | Where-Object { $_ -notin $updatedMembers.SamAccountName }
            if ($missingUsers.Count -gt 0) {
                throw "�����û����ʧ�ܣ�$($missingUsers -join ', ')"
            }
        } -ArgumentList $validUserSams, $groupSam -ErrorAction Stop

        # �����ɹ���ʾ
        $successMsg = if ($validUsers.Count -eq 1) {
            "�û�[$($validUsers[0].DisplayName)]�ѳɹ�������[$groupName]"
        } else {
            $successUserList = @()
            foreach ($user in $validUsers) {
                $successUserList += "$($user.DisplayName)��$($user.SamAccountName)��"
            }
            "�ѳɹ���� $($validUsers.Count) ���û�����[$groupName]��`n`n$($successUserList -join "`n")"
        }
        [System.Windows.Forms.MessageBox]::Show($successMsg, "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        LoadUserList  # ˢ���û��б������û���������Ϣ��
        $script:connectionStatus = "�����ӵ����: $($script:comboDomain.SelectedItem.Name)��Զ��ִ�У�"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "���ʧ��: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("����û�����ʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


function ModifyGroup {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    if ($script:groupDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("��ѡ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    # 1. ��ȡѡ����Ļ�����Ϣ����DataGridView��
    $groupRow = $script:groupDataGridView.SelectedRows[0]
    $groupData = $groupRow.DataBoundItem
    $originalSam = $groupData.SamAccountName.ToString().Trim()
    
    if ([string]::IsNullOrEmpty($originalSam)) {
        [System.Windows.Forms.MessageBox]::Show("ѡ������˺���ϢΪ�գ�������ѡ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 2. Զ�̻�ȡ����������ԣ���DN��CN��DisplayName����Ա������
    # ������$hasMembers ������Ƿ������Ա��$checkNestedMembers �����Ƿ���Ƕ�׳�Ա
    $hasMembers = $false
    $checkNestedMembers = $false  # ���������$true=����Ƕ�׳�Ա��$false=��ֱ�ӳ�Ա
    try {
        $script:connectionStatus = "���ڼ�������ϸ��Ϣ..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        $originalGroup = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($targetSam, $checkNested)
            Import-Module ActiveDirectory -ErrorAction Stop
            # 1. ��ȡ��������� + Member���ԣ����ڼ���ֱ�ӳ�Ա��
            $adGroup = Get-ADGroup -Identity $targetSam `
                        -Properties DistinguishedName, DisplayName, Description, Member `
                        -ErrorAction Stop
            
            # 2. �����Ա����������ֱ��/Ƕ�ף�
            if ($checkNested) {
                # ����Ƕ�׳�Ա���ݹ��ȡ��
                $allMembers = Get-ADGroupMember -Identity $adGroup -Recursive -ErrorAction Stop
                $memberCount = $allMembers.Count
            }
            else {
                # ��ֱ�ӳ�Ա��Member���Դ洢ֱ�ӳ�ԱDN������Ϊ0��
                $memberCount = if ($adGroup.Member -and $adGroup.Member.Count -gt 0) { $adGroup.Member.Count } else { 0 }
            }

            # 3. ���������� + ��Ա���������ں����ж��Ƿ���Ҫˢ���û��б�
            return [PSCustomObject]@{
                Name              = $adGroup.Name
                DistinguishedName = $adGroup.DistinguishedName
                DisplayName       = $adGroup.DisplayName
                Description       = $adGroup.Description
                MemberCount       = $memberCount  # ��������Ա����
            }
        } -ArgumentList $originalSam, $checkNestedMembers -ErrorAction Stop

        # ��ȡԶ�̷��صĹؼ����ԣ����ػ��棩
        $originalCN = $originalGroup.Name          
        $originalDN = $originalGroup.DistinguishedName  
        $originalDisplayName = $originalGroup.DisplayName
        $originalDescription = $originalGroup.Description
        $originalMemberCount = $originalGroup.MemberCount  # ��������¼ԭʼ��Ա����
        
        # ������������Ƿ������Ա����Ա�� > 0 ��Ϊ$true��
        if ($originalMemberCount -gt 0) {
            $hasMembers = $true
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "��������Ϣʧ��: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("��������ϸ��Ϣʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # 3. ��ȡ�û��������ֵ�������ظ�ʽУ�飨ԭ�߼����䣩
    $newCN = $script:textGroupName.Text.Trim()          
    $newSam = $script:textGroupSamAccount.Text.Trim()   
    $newDisplayName = $newCN                            
    $newDesc = $script:textGroupDescription.Text.Trim() 

    # 3.1 ������У��
    if (-not ($newCN -and $newSam)) {
        [System.Windows.Forms.MessageBox]::Show("�����ƣ�CN�������˺�Ϊ������", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 3.2 CN��ʽУ�飨��ֹLDAP�����ַ���
    if ($newCN -match '[,=\+<>;#\"\\]') {
        [System.Windows.Forms.MessageBox]::Show('�����Ʋ��ܰ������������ַ���, = + < > ; # " \', "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 3.3 SamAccountName��ʽУ�飨AD�������ƣ�
    if ($newSam -match "[^\w\-]") {
        [System.Windows.Forms.MessageBox]::Show("���˺Ų��ܰ��������ַ���������ĸ�����֡��»��ߡ����ַ���", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ($newSam.Length -gt 20) {
        [System.Windows.Forms.MessageBox]::Show("���˺ų��Ȳ��ܳ���20���ַ�", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 4. Զ�̼����SamAccountName��Ψһ�ԣ���Sam���޸ģ�ԭ�߼����䣩
    if ($newSam -ne $originalSam) {
        try {
            $script:connectionStatus = "���ڼ�������˺ſ�����..."
            UpdateStatusBar
            $script:mainForm.Refresh()

            $samExists = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($newSamAccount)
                Import-Module ActiveDirectory -ErrorAction Stop
                $existing = Get-ADGroup -Filter "SamAccountName -eq '$newSamAccount'" -ErrorAction SilentlyContinue
                return $null -ne $existing
            } -ArgumentList $newSam -ErrorAction Stop

            if ($samExists) {
                [System.Windows.Forms.MessageBox]::Show("�����˺�[$newSam]�Ѵ��ڣ������", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            $script:connectionStatus = "����˺ſ�����ʧ��: $errorMsg"
            UpdateStatusBar
            $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
            [System.Windows.Forms.MessageBox]::Show("��������˺�ʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }

    # 5. ����Ƿ���ʵ���޸ģ����޸���ֱ�ӷ��أ�ԭ�߼����䣩
    if ($newCN -eq $originalCN -and $newSam -eq $originalSam -and $newDisplayName -eq $originalDisplayName -and $newDesc -eq $originalDescription) {
        [System.Windows.Forms.MessageBox]::Show("δ��⵽�κ��޸�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 6. ȷ���޸Ĳ�����ԭ�߼����䣬���Ż�������ֵ��ʾ��
    $displayNewDesc = if ([string]::IsNullOrEmpty($newDesc)) { "��" } else { $newDesc }
    $confirmMsg = "ȷ���޸��顾$originalCN��$originalSam������`n"
    $confirmMsg += "ע�⣺�޸������ƣ�CN����ı�����AD�е�Ŀ¼·����DN��`n`n"
    $confirmMsg += "�����ƣ�CN����$newCN`n"
    $confirmMsg += "���˺ţ�Sam����$newSam`n"
    $confirmMsg += "��������$displayNewDesc"

    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, "ȷ���޸���", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne 'Yes') {
        return
    }

    # 7. Զ��ִ���޸Ĳ�����������+���Ը��£�ԭ�߼����䣩
    try {
        $script:connectionStatus = "����Զ���޸��顾$originalCN��..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        $modifiedGroup = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($origSam, $origDN, $origCN, $newCN, $newSam, $newDisplayName, $newDesc)
            Import-Module ActiveDirectory -ErrorAction Stop

            # ���޸������ƣ�CN������������AD����
            if ($newCN -ne $origCN) {
                $newDN = $origDN -replace "^CN=$([regex]::Escape($origCN)),", "CN=$newCN,"
                Rename-ADObject -Identity $origDN `
                               -NewName $newCN `
                               -ErrorAction Stop
            }

            # ���������ԣ�SamAccountName��DisplayName��Description��
            Set-ADGroup -Identity $origSam `
                        -SamAccountName $newSam `
                        -DisplayName $newDisplayName `
                        -Description $newDesc `
                        -ErrorAction Stop

            # ��֤�޸Ľ����������������
            $verifyIdentity = if ($newSam -eq $origSam) { $origSam } else { $newSam }
            return Get-ADGroup -Identity $verifyIdentity `
                               -Properties DistinguishedName, DisplayName, Description `
                               -ErrorAction Stop
        } -ArgumentList $originalSam, $originalDN, $originalCN, $newCN, $newSam, $newDisplayName, $newDesc -ErrorAction Stop

        # 8. ������֤����ʾ�ɹ���ԭ�߼����䣩
        if ($modifiedGroup.Name -ne $newCN -or $modifiedGroup.SamAccountName -ne $newSam) {
            throw "Զ���޸�ִ�гɹ��������ص����Բ�ƥ��"
        }
        $displayModifiedDesc = if ([string]::IsNullOrEmpty($modifiedGroup.Description)) { "��" } else { $modifiedGroup.Description }
        [System.Windows.Forms.MessageBox]::Show("���޸ĳɹ�`n`n" +
            "�����ƣ�CN����$($modifiedGroup.Name)`n" +
            "���˺ţ�Sam����$($modifiedGroup.SamAccountName)`n" +
            "��ʾ���ƣ�$($modifiedGroup.DisplayName)`n" +
            "������$displayModifiedDesc", 
            "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # 9. ����������ˢ���û��б����������Ա����ˢ���û��ġ������顱��Ϣ��
        LoadGroupList  # ʼ��ˢ�����б�չʾ�޸ĺ������Ϣ��
        if ($hasMembers) {
            LoadUserList   # �����������Աʱ��ˢ���û��б�ͬ���û������������ݣ�
            $refreshTip = "����ͬ��ˢ���û��б�"
        } else {
            $refreshTip = "��δˢ���û��б�������Ա��"
        }

        # 10. ����״̬��������ˢ��״̬��ʾ��
        $script:connectionStatus = "�����ӵ����: $($script:comboDomain.SelectedItem.Name)��Զ��ִ�У� $refreshTip"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "�޸���ʧ��: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("���޸�ʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


function DeleteGroup {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ���ÿ��أ��Ƿ���Ƕ�׳�Ա���������ã����������Ӱ�����ܣ�
    $checkNestedMembers = $false  # $true=����Ƕ�׳�Ա��$false=��ֱ�ӳ�Ա

    # 1. ����Ƿ�ѡ���飨֧��Ctrl��ѡ��
    if ($script:groupDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("��ѡ����Ҫɾ������", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 2. ��ȡѡ������Ϣ + ��׼����Ա�����������޸���
    $selectedGroups = @()
    $script:connectionStatus = "���ڼ��ѡ����ĳ�Ա���..."
    UpdateStatusBar
    $script:mainForm.Refresh()

    foreach ($row in $script:groupDataGridView.SelectedRows) {
        $group = $row.DataBoundItem
        if (-not $group) { continue }

        # ������Ϣ��ȡ
        $groupSam = if ($group.SamAccountName -ne $null) { $group.SamAccountName.ToString().Trim() } else { "" }
        $groupName = if ($group.Name -ne $null) { $group.Name.ToString().Trim() } else { "δ֪��" }

        if ([string]::IsNullOrEmpty($groupSam)) {
            Write-Warning "������Ч�飨�˺�Ϊ�գ���$groupName"
            continue
        }

        # �������޸���ʹ��Get-ADGroup��Member���Լ���Ա�����ɿ���
        try {
            $memberCount = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($samAccountName, $checkNested)
                Import-Module ActiveDirectory -ErrorAction Stop

                # ��һ������ȡ����󣨻�ȡMember���ԣ�
                $adGroup = Get-ADGroup -Identity $samAccountName -Properties Member -ErrorAction Stop

                # �ڶ����������Ա����������ֱ��/Ƕ�ף�
                if ($checkNested) {
                    # ����Ƕ�׳�Ա���ݹ��ȡ�����û�/�������
                    $allMembers = Get-ADGroupMember -Identity $adGroup -Recursive -ErrorAction Stop
                    return $allMembers.Count
                }
                else {
                    # ��ֱ�ӳ�Ա��Member���԰�������ֱ�ӳ�Ա��DN��
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
            # �鲻����
            $memberCount = -2
            Write-Warning "�����[$groupName]ʧ�ܣ��鲻����"
        }
        catch [System.UnauthorizedAccessException] {
            # Ȩ�޲���
            $memberCount = -3
            Write-Warning "�����[$groupName]��Աʧ�ܣ�Ȩ�޲���"
        }
        catch {
            # ��������
            $memberCount = -1
            Write-Warning "�����[$groupName]��Աʧ�ܣ�$($_.Exception.Message)"
        }

        # ����ѡ���б�������ϸ״̬��
        $selectedGroups += [PSCustomObject]@{
            SamAccountName = $groupSam
            Name           = $groupName
            MemberCount    = $memberCount  # -3=Ȩ�޲��㣻-2=�鲻���ڣ�-1=��������>=0=��Ա��
            CheckType      = if ($checkNestedMembers) { "����Ƕ��" } else { "��ֱ��" }
        }
    }

    # 3. ������Ч�鲢��ʾ
    #$validGroups = $selectedGroups | Where-Object { $_.MemberCount -ge 0 }
    #$invalidGroups = $selectedGroups | Where-Object { $_.MemberCount -lt 0 }
	$validGroups = @($selectedGroups | Where-Object { $_.MemberCount -ge 0 })
	$invalidGroups = @($selectedGroups | Where-Object { $_.MemberCount -lt 0 })

    if ($invalidGroups.Count -gt 0) {
        $invalidMsg = "�������޷�ɾ����`n"
        foreach ($g in $invalidGroups) {
            switch ($g.MemberCount) {
                -3 { $invalidMsg += "- $($g.Name)��Ȩ�޲��㣬�޷����ʳ�Ա��Ϣ`n" }
                -2 { $invalidMsg += "- $($g.Name)���鲻����`n" }
                default { $invalidMsg += "- $($g.Name)����Ա���ʧ��`n" }
            }
        }
        [System.Windows.Forms.MessageBox]::Show($invalidMsg, "�޷��������", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }

    if ($validGroups.Count -eq 0) {
        $script:connectionStatus = "�����ӵ����: $($script:comboDomain.SelectedItem.Name)��Զ��ִ�У�"
        UpdateStatusBar
        return
    }

    # 4. ����ȷ����ʾ�����ص��޸����Ż���ѡ/��ѡ�жϣ�ȷ��������ʾ׼ȷ��
	$groupInfoLines = @()
	foreach ($g in $validGroups) {
		if ($g.MemberCount -gt 0) {
			$groupInfoLines += "$($g.Name)��$($g.SamAccountName)��- $($g.CheckType)��Ա��$($g.MemberCount)��"
		}
		else {
			$groupInfoLines += "$($g.Name)��$($g.SamAccountName)��- �޳�Ա"
		}
	}
	$groupListText = $groupInfoLines -join "`n"

    # ���޸����ġ���ȷ���ֵ�ѡ/��ѡ���������������ֵ
	if ($validGroups.Count -eq 1) {
		$confirmMsg = "ȷ������ɾ����������`n`n$groupListText`n`nע�⣺�����Գ�Աֱ��ɾ����"
	}
	else {
		$confirmMsg = "��ѡ�� $($validGroups.Count) ����Ч�飬ȷ������ɾ����`n`n$groupListText`n`nע�⣺�����Գ�Աֱ��ɾ����"
	}

    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, "ȷ��ɾ����", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -ne 'Yes') {
        $script:connectionStatus = "�����ӵ����: $($script:comboDomain.SelectedItem.Name)��Զ��ִ�У�"
        UpdateStatusBar
        return
    }

    # 5. ����ɾ���� + ����Ƿ���Ҫˢ���û��б�
    $successCount = 0
    $failedGroups = @()
    $hasDeletedGroupWithMembers = $false
    $script:connectionStatus = "����ɾ�� $($validGroups.Count) ����..."
    UpdateStatusBar
    $script:mainForm.Refresh()

    foreach ($g in $validGroups) {
        try {
            # Զ��ִ��ɾ��
            Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($groupSam)
                Import-Module ActiveDirectory -ErrorAction Stop
                # ǿ��ɾ������ʹ�г�Ա��
                Remove-ADGroup -Identity $groupSam -Confirm:$false -ErrorAction Stop
                # ��֤ɾ��
                $exists = Get-ADGroup -Filter "SamAccountName -eq ""$groupSam""" -ErrorAction SilentlyContinue
                if ($exists) { throw "ɾ�����Կɲ�ѯ����" }
            } -ArgumentList $g.SamAccountName -ErrorAction Stop

            $successCount++
            # ��ɾ�������г�Ա���飬���ˢ��
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

    # 6. ��ʾɾ�����
    $resultMsg = if ($successCount -eq $validGroups.Count) {
        if ($successCount -eq 1) {
            "��[$($validGroups[0].Name)]������ɾ��"
        }
        else {
            $successList = $validGroups.Name -join "��"
            "�ѳɹ�ɾ�� $successCount ���飺`n$successList"
        }
    }
    else {
        $successNames = $validGroups | Where-Object { $_.SamAccountName -notin $failedGroups.SamAccountName } | Select-Object -ExpandProperty Name
        $successList = $successNames -join "��"
        $failedList = $failedGroups | ForEach-Object { "$($_.Name)������$($_.Error)��" } -join "`n"
        "ɾ����ɣ�`n�ɹ���$successCount ���飨$successList��`nʧ�ܣ�$($failedGroups.Count) ���飺`n$failedList"
    }

    $msgIcon = if ($failedGroups.Count -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
    [System.Windows.Forms.MessageBox]::Show($resultMsg, "ɾ�����", [System.Windows.Forms.MessageBoxButtons]::OK, $msgIcon)

    # 7. ����ˢ���б�
    $script:groupDataGridView.ClearSelection()
    LoadGroupList  # ʼ��ˢ�����б�

    if ($hasDeletedGroupWithMembers) {
        LoadUserList   # ��ɾ���г�Ա����ʱˢ���û��б�
        $refreshTip = "����ͬ��ˢ���û��б�"
    }
    else {
        $refreshTip = "��δˢ���û��б�ɾ���ľ�Ϊ�޳�Ա�飩"
    }

    # 8. ����״̬��
    $statusText = if ($failedGroups.Count -eq 0) {
        "�ѳɹ�ɾ�� $successCount ���� $refreshTip"
    }
    else {
        "ɾ����ɣ��ɹ� $successCount ����ʧ�� $($failedGroups.Count) �� $refreshTip"
    }
    $script:connectionStatus = $statusText
    UpdateStatusBar
    $script:statusOutputLabel.ForeColor = if ($failedGroups.Count -eq 0) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkOrange }
}




function RemoveUserFromGroup {
    # 1. ����������
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ֧�ֶ�ѡ�û������ѡ���û�����
    if ($script:userDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("����ѡ����Ҫ�������Ƴ����û���֧��Ctrl��ѡ��", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ���ѡ��Ŀ���飨�����飩
    if ($script:groupDataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("����ѡ��Ҫ�Ƴ��û���Ŀ���飨��֧�ֵ����飩", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $selectedGroup = $script:groupDataGridView.SelectedRows[0].DataBoundItem
    if (-not $selectedGroup) {
        [System.Windows.Forms.MessageBox]::Show("ѡ���������쳣��������ѡ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ��ȡ����ѡ���û��ĺ�����Ϣ
    $selectedUsers = @()
    foreach ($userRow in $script:userDataGridView.SelectedRows) {
        # ��ȡ�û�SamAccountName������ǿգ�
        if ($userRow.Cells["SamAccountName"].Value -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("ѡ���û��������쳣���˺�Ϊ�գ���������ѡ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $username = $userRow.Cells["SamAccountName"].Value.ToString().Trim()

        # ��ȡ�û���ʾ����Ϊ�������˺Ŵ��棩
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

    # ��ȡĿ������Ϣ
    if ($selectedGroup.SamAccountName -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("ѡ������˺���ϢΪ�գ�������ѡ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $groupSam = $selectedGroup.SamAccountName.ToString().Trim()
    $groupName = if ($selectedGroup.Name -ne $null) {
        $selectedGroup.Name.ToString().Trim()
    } else {
        $groupSam
    }

    # ��������û��Ƿ�������
    $nonMemberUsers = @()  # �������е��û�
    $validUsers = @()     # ��ɾ������Ч�û�
    try {
        $script:connectionStatus = "���ڼ�� $($selectedUsers.Count) ���û������Ա��ϵ..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # Զ��������ȡ�������г�Ա������AD���ã�
        $groupMembers = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($targetGroup)
            Import-Module ActiveDirectory -ErrorAction Stop
            $members = Get-ADGroupMember -Identity $targetGroup -Recursive -ErrorAction Stop
            return $members.SamAccountName
        } -ArgumentList $groupSam -ErrorAction Stop

        # �Ա�ɸѡ����������û� vs ��ɾ�����û�
        foreach ($user in $selectedUsers) {
            if ($groupMembers -notcontains $user.SamAccountName) {
                $nonMemberUsers += $user
            } else {
                $validUsers += $user
            }
        }

        # ��ʾ��������û������жϲ�����
        if ($nonMemberUsers.Count -gt 0) {
            $nonMemberNames = @()
            foreach ($user in $nonMemberUsers) {
                $nonMemberNames += "$($user.DisplayName)��$($user.SamAccountName)��"
            }
            [System.Windows.Forms.MessageBox]::Show("�����û�������[$groupName]�У������Ƴ���`n`n$($nonMemberNames -join "`n")", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }

        # �������û��������飬ֱ���˳�
        if ($validUsers.Count -eq 0) {
            $script:connectionStatus = "�����ӵ����: $($script:comboDomain.SelectedItem.Name)��Զ��ִ�У�"
            UpdateStatusBar
            return
        }
    }
    catch {
        $script:connectionStatus = "����Ա��ϵʧ��: $($_.Exception.Message)"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("����û�-���ϵʧ�ܣ�`n$($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ����ȷ��ɾ������
    $validUserNames = @()
    foreach ($user in $validUsers) {
        $validUserNames += "$($user.DisplayName)��$($user.SamAccountName)��"
    }
    
    $confirmTitle = if ($validUsers.Count -eq 1) { "ȷ���Ƴ��û�" } else { "ȷ�������Ƴ��û�" }
    $confirmMsg = if ($validUsers.Count -eq 1) {
        "ȷ�����û�`n`n$($validUserNames -join "`n")`n`n����[$groupName��$groupSam��]���Ƴ���`n`n�Ƴ����û���ʧȥ�����Ȩ�ޣ�"
    } else {
        "��ѡ�� $($validUsers.Count) ���û���ȷ����������[$groupName��$groupSam��]���Ƴ���`n`n���Ƴ��û���`n$($validUserNames -join "`n")`n`n�Ƴ����û���ʧȥ�����Ȩ�ޣ�"
    }

    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, $confirmTitle, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -ne 'Yes') {
        return
    }

    # ����Զ��ɾ���û�
    try {
        $script:connectionStatus = "����Զ���Ƴ� $($validUsers.Count) ���û�from��[$groupName]..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # ��ȡ��ɾ���û���Sam�˺��б�
        $validUserSams = $validUsers.SamAccountName

        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($users, $group)
            Import-Module ActiveDirectory -ErrorAction Stop
            # �����Ƴ�
            Remove-ADGroupMember -Identity $group -Members $users -Confirm:$false -ErrorAction Stop
            
            # ��֤��ȷ�������û������Ƴ�
            $remainingMembers = Get-ADGroupMember -Identity $group -Recursive -ErrorAction Stop
            $remainingUsers = $users | Where-Object { $_ -in $remainingMembers.SamAccountName }
            if ($remainingUsers.Count -gt 0) {
                throw "�����û��Ƴ�ʧ�ܣ�$($remainingUsers -join ', ')"
            }
        } -ArgumentList $validUserSams, $groupSam -ErrorAction Stop

        # �����ɹ���ʾ
        $successMsg = if ($validUsers.Count -eq 1) {
            "�û�[$($validUsers[0].DisplayName)]�ѳɹ�����[$groupName]���Ƴ�"
        } else {
            $successUserList = @()
            foreach ($user in $validUsers) {
                $successUserList += "$($user.DisplayName)��$($user.SamAccountName)��"
            }
            "�ѳɹ�����[$groupName]���Ƴ� $($validUsers.Count) ���û���`n`n$($successUserList -join "`n")"
        }
        [System.Windows.Forms.MessageBox]::Show($successMsg, "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        LoadUserList  # ˢ���û��б������û���������Ϣ��
        $script:connectionStatus = "�����ӵ����: $($script:comboDomain.SelectedItem.Name)��Զ��ִ�У�"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "�Ƴ��û�ʧ��: $errorMsg"
        UpdateStatusBar
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show("�Ƴ��û�from��ʧ�ܣ�`n$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

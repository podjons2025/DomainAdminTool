<# 
OU�������ĺ���
#>

function LoadOUList {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    try {
        $script:connectionStatus = "���ڴ���ض�ȡOU�б�..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # Զ�̻�ȡ����OU����������µ�OU��
        $script:allOUs = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            Import-Module ActiveDirectory -ErrorAction Stop
            Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName |
                Where-Object { $_.Name -ne "Domain Controllers" } |			
                Select-Object Name, DistinguishedName |
                Sort-Object Name
        } -ErrorAction Stop

        # ����OU��νṹ�����ɴ���ε���ʾ���ƣ�֧������µ�OU��
        $script:allOUs = $script:allOUs | ForEach-Object {
            $dn = $_.DistinguishedName
            $ouParts = @()
            # ��ȡDN�е�����OU���������DC���֣�
            $dn -split ',' | ForEach-Object {
                if ($_ -match '^OU=(.+)') {
                    $ouParts += $matches[1]
                }
            }
            # ��תOU���˳��DN���Ǵ��ӵ�������ת��Ϊ�Ӹ����ӣ�
            $hierarchyParts = $ouParts[($ouParts.Count - 1)..0]
            $displayHierarchy = if ($hierarchyParts.Count -gt 1) {
                $hierarchyParts -join ' > '  # ��㼶�ü�ͷ����
            }
            else {
                $_.Name  # ����OU������µ�OU��ֱ����ʾ����
            }
            [PSCustomObject]@{
                Name              = $_.Name
                DistinguishedName = $_.DistinguishedName
                DisplayHierarchy  = $displayHierarchy  # ��λ���ʾ����
            }
        } | Sort-Object DisplayHierarchy  # ����νṹ����

        return $script:allOUs
    }
    catch {
        $errorMsg = $_.Exception.Message
        $script:connectionStatus = "��ȡOU�б�ʧ�ܣ�$errorMsg"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("��ȡOU�б�ʧ�ܣ�$errorMsg", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
}

# �л�OU��֯��֧�������Users������
function SwitchOU {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ����OU�б�����νṹ��
    $ous = LoadOUList
    if (-not $ous -or $ous.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("δ�ҵ��κ�OU��֯", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ��ȡĬ��Users����������Ϣ���޸���������߼���
    $defaultUsersOU = $null
    $domainDN = $null
    
    # ���ȴ�domainContext��ȡ���
    if ($script:domainContext -and $script:domainContext.DomainInfo) {
        $domainDN = $script:domainContext.DomainInfo.DefaultPartition
    }
    # �ӵ�ǰOU��ȡ������ؼ��޸���֧�ֶ����OU��
    if (-not $domainDN -and $script:currentOU) {
        # ����ƥ��DN����������DC�����������������"OU=��,OU=��,DC=domain,DC=com"����ȡ"DC=domain,DC=com"
        if ($script:currentOU -match '(DC=.+)$') {
            $domainDN = $matches[1]
        }
    }
    
    # ������ȷ��Users����·��
    if ($domainDN) {
        $defaultUsersOU = "CN=Users,$domainDN"
        $script:allUsersOU = $defaultUsersOU  # ͳһUsers����·���������ظ�����
    }

    # �����̶�ѡ����������Ĭ��Users�����Ƴ�����Users��
    $fixedItems = @()
    # ������ѡ��
    if ($domainDN) {
        $fixedItems += [PSCustomObject]@{
            Name              = "���"
            DistinguishedName = $domainDN
            DisplayHierarchy  = "��� ($($domainDN -replace 'DC=','.' -replace ',',''))"
        }
    }
    # ���Users����ѡ���ʾΪ"Ĭ��(CN=Users,DC=bocmodc3,DC=com)"��ʽ��
    if (-not [string]::IsNullOrWhiteSpace($defaultUsersOU)) {
        $fixedItems += [PSCustomObject]@{
            Name              = "Ĭ��Users"
            DistinguishedName = $defaultUsersOU
            DisplayHierarchy  = "Ĭ��($defaultUsersOU)"  # �ؼ��޸ģ���ʾ����DN
        }
    }

    # �ϲ��̶�ѡ��Ͳ�λ�OU�б�
    $displayItems = $fixedItems + $ous

    # ����OUѡ��Ի���
    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "ѡ��OU��֯"
    $ouForm.Size = New-Object System.Drawing.Size(500, 350)  # �ӿ�������ʾ����·��
    $ouForm.StartPosition = "CenterScreen"
    $ouForm.MaximizeBox = $false
    $ouForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

    # ��ť���
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = "Bottom"
    $buttonPanel.Height = 40
    $buttonPanel.Padding = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
    $buttonPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

    # �б����ʾ��νṹ��
    $ouListBox = New-Object System.Windows.Forms.ListBox
    $ouListBox.Dock = "Fill"
    $ouListBox.DisplayMember = "DisplayHierarchy"  # ��ʾ��λ�����
    $ouListBox.ValueMember = "DistinguishedName"
    $ouListBox.Items.AddRange($displayItems)
    $ouListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ouListBox.Font = New-Object System.Drawing.Font("΢���ź�", 9)  # ��������
    if ($script:currentOU) {
        $selectedItem = $ouListBox.Items | Where-Object { $_.DistinguishedName -eq $script:currentOU }
        if ($selectedItem) {
            $ouListBox.SelectedItem = $selectedItem
        }
    }

    # ȷ����ť
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "ȷ��"
    $okButton.Width = 100
    $okButton.Height = 30
    $okButton.FlatAppearance.BorderSize = 1	
    $okButton.Location = New-Object System.Drawing.Point(130, 5)
    $okButton.Add_Click({
        if ($ouListBox.SelectedItem) {
            $selectedItem = $ouListBox.SelectedItem
            $script:currentOU = $selectedItem.DistinguishedName
            $script:textOU.Text = $script:currentOU
            
            # �Ƴ�����Users����߼�����ѡ����ɾ����
            $script:allUsersOU = $null

            # ���������
            $script:textSearch.Text = ""
            $script:textGroupSearch.Text = ""
            
            $ouForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        }
    })

    # ȡ����ť
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "ȡ��"
    $cancelButton.Width = 100
    $cancelButton.Height = 30
    $cancelButton.FlatAppearance.BorderSize = 1
    $cancelButton.Location = New-Object System.Drawing.Point(255, 5)
    $cancelButton.Add_Click({
        $ouForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    })

    # ��ӿؼ�
    $buttonPanel.Controls.Add($okButton)
    $buttonPanel.Controls.Add($cancelButton)
    $ouForm.Controls.Add($ouListBox)
    $ouForm.Controls.Add($buttonPanel)

    if ($ouForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        # ��ȡ��ʾ���ƣ�����Σ�
        $displayName = if ($selectedItem.DisplayHierarchy) {
            $selectedItem.DisplayHierarchy
        } else {
            $script:currentOU.Split(',')[0] -replace 'CN=', ''
        }
        $script:connectionStatus = "���л���OU��$displayName"
        UpdateStatusBar
        
        # ˢ���б�
        LoadUserList
        LoadGroupList
    }
}

# �½�OU��֯
function CreateNewOU {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("δѡ��ǰOU�������л���Ŀ��OU���ٲ���", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $newOUName = [Microsoft.VisualBasic.Interaction]::InputBox(
		"��������OU�����ƣ�ʾ����ITDepartment�����񲿣�`n`nע�⣺���ɰ��������ַ���/\=+:*#$@?!~`"<>|��", 
		"�½�OU��֯", 
		""
    )

    # ����У��
	if ([string]::IsNullOrEmpty($newOUName)) {
		return
	}
	elseif ([string]::IsNullOrWhiteSpace($newOUName)) {
		[System.Windows.Forms.MessageBox]::Show("OU���Ʋ���Ϊ�ջ�������ո�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
		return
	}
		
    # �����ַ�У��
    $invalidChars = '[\\/=+:*#$@?!~"<>|]'
    if ($newOUName -match $invalidChars) {
        $matchedChar = $matches[0]
        [System.Windows.Forms.MessageBox]::Show("OU���ư����Ƿ��ַ���`"$matchedChar`"`n��ɾ�������ԣ�", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ȷ��������
    $parentDN = $null
    $parentDisplay = $null

    # ���⴦����ǰOU��CN=Users,DC=bocmodc3,DC=comʱ��������Ϊ����DC=bocmodc3,DC=com
    if ($script:currentOU -eq "CN=Users,DC=bocmodc3,DC=com") {
        $parentDN = "DC=bocmodc3,DC=com"
        $parentDisplay = "��� (bocmodc3.com)"
    }
    # ���������ֱ���Ե�ǰOU��Ϊ������
    else {
        $parentDN = $script:currentOU
        
        # ���ɸ������Ѻ���ʾ����
        $ouParts = @()
        $parentDN -split ',' | ForEach-Object {
            if ($_ -match '^OU=(.+)') { $ouParts += $matches[1] }
            elseif ($_ -match '^CN=Users') { $ouParts += "Users����" }
            elseif ($_ -match '^DC=.+') { $ouParts += "���" }
        }
        $parentDisplay = if ($ouParts.Count -gt 0) {
            $ouParts[($ouParts.Count - 1)..0] -join ' > '
        } else {
            $parentDN.Split(',')[0] -replace '^(OU|CN)=', ''
        }
    }

    # ������OU����·��
    $newOUFullDN = "OU=$newOUName,$parentDN"
    
    # ȷ�ϴ���
    $confirmMsg = "ȷ�ϴ���OU��`n������: $parentDisplay`n��OU����: $newOUName`n����·��: $newOUFullDN`n`nע�⣺�˲�������ѡ���ĸ������´���OU"
    $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "ȷ���½�", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $script:connectionStatus = "���ڴ���OU��$newOUName..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # Զ�̴���OU
        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($ouName, $parentDN)
            Import-Module ActiveDirectory -ErrorAction Stop
            New-ADOrganizationalUnit -Name $ouName -Path $parentDN -ProtectedFromAccidentalDeletion $true -ErrorAction Stop
        } -ArgumentList $newOUName, $parentDN -ErrorAction Stop

        # �����ӳ�ȷ��AD�������ͬ��
        Start-Sleep -Milliseconds 500

        # �����ɹ�����
        $script:currentOU = $newOUFullDN
        $script:textOU.Text = $script:currentOU

        $script:connectionStatus = "OU�����ɹ���$newOUFullDN"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU�����ɹ���`n����·����$newOUFullDN", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # ǿ��ˢ��OU�б�
        LoadOUList | Out-Null

        # ��UI�߳���ˢ���û������б�
        if ($script:mainForm.InvokeRequired) {
            # ��ȷָ��ί������ΪAction
            $script:mainForm.Invoke([System.Action]{
                try {
                    LoadUserList
                    LoadGroupList
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("ˢ���б�ʧ��: $($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            })
        }
        else {
            try {
                LoadUserList
                LoadGroupList
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("ˢ���б�ʧ��: $($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }

    } catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "already exists") {
            $errorMsg = "OU�Ѵ��ڣ���������ƣ��磺$newOUName_2��"
        } elseif ($errorMsg -match "permission") {
            $errorMsg = "Ȩ�޲��㣡��ȷ�Ϲ���Ա�˺�ӵ���ڸ������´���OU��Ȩ��"
        } elseif ($errorMsg -match "Path") {
            $errorMsg = "������·������$parentDN������������Ƿ����"
        } elseif ($errorMsg -match "invalid DN syntax") {
            $errorMsg = "·���﷨����$newOUFullDN�����������Ƿ���������ַ�"
        }
        $script:connectionStatus = "OU����ʧ�ܣ�$errorMsg"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


# ������OU��֯
function RenameExistingOU {
    [CmdletBinding()]
    param()

    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("δ��ȡ����ǰOU��Ϣ��������������أ�", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ���岻������������ϵͳ�ؼ�����
    $protectedContainers = @(
        "CN=Users,DC=bocmodc3,DC=com",  # Ĭ��Users����
        "DC=bocmodc3,DC=com"            # �������
    )

    # ��鵱ǰOU�Ƿ�Ϊϵͳ�ؼ�����
    if ($protectedContainers -contains $script:currentOU) {
        $containerName = if ($script:currentOU -eq "DC=bocmodc3,DC=com") {
            "���������$($script:currentOU -replace 'DC=','.' -replace ',','')��"
        } else {
            "Ĭ��Users������CN=Users��"
        }
        [System.Windows.Forms.MessageBox]::Show(
            "$containerName ��ϵͳ�ؼ���������������������", 
            "������ֹ", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Stop
        )
        return
    }

    # ������ǰOU�����ƺ͸�·��
    if ($script:currentOU -match '^OU=(.+?),(.+)') {
        $currentOUName = $matches[1]
        $parentDN = $matches[2]
    } else {
        [System.Windows.Forms.MessageBox]::Show("��ǰѡ�еĲ�����Ч��OU����", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ��ʾ��ǰOU��Ϣ
    $ouParts = @()
    $script:currentOU -split ',' | ForEach-Object {
        if ($_ -match '^OU=(.+)') { $ouParts += $matches[1] }
    }
    $displayHierarchy = if ($ouParts.Count -gt 0) {
        $ouParts[($ouParts.Count - 1)..0] -join ' > '
    } else {
        $currentOUName
    }

    # ��ȡ������
    $newOUName = [Microsoft.VisualBasic.Interaction]::InputBox(
        "�������µ�OU����`n`n��ǰOU��$displayHierarchy`n`n��ǰ���ƣ�$currentOUName`n`nע�⣺���ɰ��������ַ���/\=+:*#$@?!~`"<>|��", 
        "������OU��֯", 
        ""
    )

    # ����У��
    if ([string]::IsNullOrEmpty($newOUName)) {
        return
    }
    elseif ([string]::IsNullOrWhiteSpace($newOUName)) {
        [System.Windows.Forms.MessageBox]::Show("OU���Ʋ���Ϊ�ջ�������ո�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    # ����δ���
    elseif ($newOUName -eq $currentOUName) {
        [System.Windows.Forms.MessageBox]::Show("�������뵱ǰ������ͬ�������޸ģ�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # �����ַ�У��
    $invalidChars = '[\\/=+:*#$@?!~"<>|]'
    if ($newOUName -match $invalidChars) {
        $matchedChar = $matches[0]
        [System.Windows.Forms.MessageBox]::Show("OU���ư����Ƿ��ַ���`"$matchedChar`"`n��ɾ�������ԣ�", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ȷ��������
    $newOUFullDN = "OU=$newOUName,$parentDN"
    $confirmMsg = "ȷ��������OU��`n��ǰOU��$displayHierarchy`nԭ���ƣ�$currentOUName`n�����ƣ�$newOUName`n������·����$newOUFullDN"
    $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "ȷ��������", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $script:connectionStatus = "����������OU��$currentOUName -> $newOUName..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # Զ��ִ��������
        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($oldDN, $newName)
            Import-Module ActiveDirectory -ErrorAction Stop
            # ʹ��Rename-ADObject������OU
            Rename-ADObject -Identity $oldDN -NewName $newName -ErrorAction Stop
        } -ArgumentList $script:currentOU, $newOUName -ErrorAction Stop

        # �������ɹ�����
        $script:currentOU = $newOUFullDN  # ���µ�ǰOU·��
        $script:textOU.Text = $script:currentOU
        
        $script:connectionStatus = "OU�������ɹ���$currentOUName -> $newOUName"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU�������ɹ���`n������·����$newOUFullDN", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
		Start-Sleep -Milliseconds 500
        # ˢ������б�
        LoadOUList
        if ($script:mainForm.InvokeRequired) {
            # ��ȷָ��ί������ΪAction
            $script:mainForm.Invoke([System.Action]{
                try {
                    LoadUserList
                    LoadGroupList
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("ˢ���б�ʧ��: $($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            })
        }
        else {
            try {
                LoadUserList
                LoadGroupList
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("ˢ���б�ʧ��: $($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }

    } catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "already exists") {
            $errorMsg = "�����Ѵ��ڣ��ø�������������Ϊ`"$newOUName`"��OU"
        } elseif ($errorMsg -match "permission") {
            $errorMsg = "Ȩ�޲��㣡��ȷ�Ϲ���Ա�˺�ӵ��OU������Ȩ��"
        } elseif ($errorMsg -match "not found") {
            $errorMsg = "OU�����ڣ������ѱ�ɾ����·������"
        } elseif ($errorMsg -match "invalid name") {
            $errorMsg = "��Ч�����Ƹ�ʽ�������Ƿ������֧�ֵ��ַ����ʽ"
        }
        
        $script:connectionStatus = "OU������ʧ�ܣ�$errorMsg"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}



# ɾ��OU��֯
function DeleteExistingOU {
    [CmdletBinding()]
    param()

    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("δ��ȡ����ǰOU��Ϣ��������������أ�", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ���岻����ɾ����ϵͳ�ؼ�����
    $protectedContainers = @(
        "CN=Users,DC=bocmodc3,DC=com",  # Ĭ��Users����
        "DC=bocmodc3,DC=com"            # �������
    )

    # ��鵱ǰOU�Ƿ�Ϊϵͳ�ؼ�����
    if ($protectedContainers -contains $script:currentOU) {
        $containerName = if ($script:currentOU -eq "DC=bocmodc3,DC=com") {
            "���������$($script:currentOU -replace 'DC=','.' -replace ',','')��"
        } else {
            "Ĭ��Users������CN=Users��"
        }
        [System.Windows.Forms.MessageBox]::Show(
            "$containerName ��ϵͳ�ؼ�������������ɾ����`n������Ļ��������ɾ���ᵼ�������쳣��", 
            "������ֹ", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Stop
        )
        return
    }

    # ��ʾOU�����Ϣ��֧������µ�OU��
    $ouParts = @()
    $script:currentOU -split ',' | ForEach-Object {
        if ($_ -match '^OU=(.+)') { $ouParts += $matches[1] }
    }
    $displayHierarchy = if ($ouParts.Count -gt 0) {
        $ouParts[($ouParts.Count - 1)..0] -join ' > '
    } else {
        $script:currentOU.Split(',')[0] -replace 'OU=',''  # ����µĶ���OU
    }
    $deleteMsg = "���棺ɾ��OU��ͬʱɾ���������ж����û����顢��OU����`n��ǰOU��Σ�$displayHierarchy`n����·����$script:currentOU`n`nȷ��Ҫɾ����"

    $confirmResult = [System.Windows.Forms.MessageBox]::Show($deleteMsg, "��Σ����ȷ��", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $script:connectionStatus = "����ɾ��OU��$displayHierarchy..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # Զ��ɾ��OU�����Ӷ���
        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($ouDN)
            Import-Module ActiveDirectory -ErrorAction Stop
            
            # �������
            try {
                Set-ADOrganizationalUnit -Identity $ouDN -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
            }
            catch {
                # ���ϵͳ������������ʾ
                if ($ouDN -in "CN=Users,DC=bocmodc3,DC=com", "DC=bocmodc3,DC=com") {
                    Write-Error "ϵͳ�ؼ�����������������"
                } else {
                    Write-Warning "�������ʱ����: $($_.Exception.Message)"
                }
            }
            
            # �ݹ�ɾ��
            Remove-ADOrganizationalUnit -Identity $ouDN -Recursive -Confirm:$false -ErrorAction Stop
        } -ArgumentList $script:currentOU -ErrorAction Stop

        # �����ӳ�ȷ��AD�������ͬ��
        Start-Sleep -Milliseconds 500

        # ɾ���ɹ�����
        $script:connectionStatus = "OUɾ���ɹ���$displayHierarchy"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU�ѳ���ɾ�����������Ӷ��󣩣�", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # ���õ�ǰOUΪ��������֧�������Users������
        $parentDN = $script:currentOU -replace '^[^,]+,', ''
        if ($parentDN -match '^OU=|^CN=Users,|^DC=') {  # ������������OU��Users���������
            $script:currentOU = $parentDN
        }
        else {
            # �����������ʱ���л������
            $script:currentOU = $parentDN
        }
        $script:textOU.Text = $script:currentOU

        # ǿ��ˢ��OU�б�
        LoadOUList | Out-Null

        # ��UI�߳���ˢ���û������б�
        if ($script:mainForm.InvokeRequired) {
            $script:mainForm.Invoke([System.Action]{
                try {
                    LoadUserList
                    LoadGroupList
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("ˢ���б�ʧ��: $($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            })
        }
        else {
            try {
                LoadUserList
                LoadGroupList
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("ˢ���б�ʧ��: $($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }

    } catch {
        $errorMsg = $_.Exception.Message
        # ���ϵͳ�ؼ������Ĵ�����ʾ�Ż�
        if ($script:currentOU -eq "CN=Users,DC=bocmodc3,DC=com") {
            $errorMsg = "�޷�ɾ��Ĭ��Users����������ϵͳ����������������ɾ����"
        } elseif ($script:currentOU -eq "DC=bocmodc3,DC=com") {
            $errorMsg = "�޷�ɾ�����������������Ļ�����ɾ���ᵼ�������������"
        } elseif ($errorMsg -match "not found") {
            $errorMsg = "OU�����ڣ������ѱ�ɾ����·������"
        } elseif ($errorMsg -match "permission") {
            $errorMsg = "Ȩ�޲��㣡��ȷ�Ϲ���Ա�˺�ӵ��OUɾ��Ȩ��"
        } elseif ($errorMsg -match "����") {
            $errorMsg = "�޷����OU���������ܸ�OU��ϵͳ��������Ҫ����Ȩ��"
        }
        
        Write-Error $errorMsg
        $script:connectionStatus = "OUɾ��ʧ�ܣ�$errorMsg"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


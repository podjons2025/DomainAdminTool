<# 
OU�������ĺ���
#>


function Get-DomainRootDN {
    # 1. ������������������������أ�
    if (-not $script:remoteSession -and -not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ���ط�����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return $null
    }

    try {
        # 2. ���ȴ�Զ�̻Ự��ȡ����Ϣ����ɿ���ֱ�Ӷ�ȡ������ã�
        $domainInfo = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            Import-Module ActiveDirectory -ErrorAction Stop
            # ��ȡ��ǰ���������Ϣ��DefaultPartition��Ϊ���DN����DC=contoso,DC=com��
            $adDomain = Get-ADDomain -ErrorAction Stop
            return @{
                DomainRootDN = $adDomain.DefaultPartition  # ���ģ����DN
                DomainDNS    = $adDomain.DNSRoot           # ��������DNS���ƣ���contoso.com��
            }
        } -ErrorAction Stop

        # 3. ��֤���������DN
        if (-not [string]::IsNullOrWhiteSpace($domainInfo.DomainRootDN)) {
            return $domainInfo.DomainRootDN
        }

        # 4. ���÷�������domainContext��ȡ�����ݾ��߼���
        if ($script:domainContext -and $script:domainContext.DomainInfo -and $script:domainContext.DomainInfo.DefaultPartition) {
            return $script:domainContext.DomainInfo.DefaultPartition
        }

        # 5. ���׷������ӵ�ǰOU�������������currentOU�Ѵ���ʱ��
        if ($script:currentOU -match '(DC=.+)$') {
            return $matches[1]
        }

        # 6. ���з���ʧ��
        [System.Windows.Forms.MessageBox]::Show("�޷���ȡ�����Ϣ���������������", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
    catch {
        $errorMsg = "��ȡ���ʧ�ܣ�$($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
}

# ����OU�б�
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

#�л�OU��֯
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

    # �ؼ��޸ģ���̬��ȡ�����Ĭ��Users����
    $domainRootDN = Get-DomainRootDN  # ����ͨ�ú�����ȡ���
    if (-not $domainRootDN) { return }  # ��ȡʧ������ֹ

    $defaultUsersOU = "CN=Users,$domainRootDN"  # ��̬����Users����·��
    $script:allUsersOU = $defaultUsersOU  # ͳһUsers����·��

    # �����̶�ѡ������Ĭ��Users��
    $fixedItems = @()
    # ������ѡ���ʾ��ʽ����� (contoso.com)��
    $domainDNS = $domainRootDN -replace 'DC=','.' -replace ',',''  # ��DC=domain,DC=comת��domain.com
    $fixedItems += [PSCustomObject]@{
        Name              = "���"
        DistinguishedName = $domainRootDN
        DisplayHierarchy  = "��� ($domainDNS)"
    }
    # ���Users����ѡ���ʾ������̬·����
    $fixedItems += [PSCustomObject]@{
        Name              = "Ĭ��Users"
        DistinguishedName = $defaultUsersOU
        DisplayHierarchy  = "Ĭ��($defaultUsersOU)"
    }

    # �ϲ��̶�ѡ��Ͳ�λ�OU�б�
    $displayItems = $fixedItems + $ous

    # ����OUѡ��Ի����߼����䣬������Դ��Ϊ��̬��
    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "ѡ��OU��֯"
    $ouForm.Size = New-Object System.Drawing.Size(500, 350)
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
    $ouListBox.DisplayMember = "DisplayHierarchy"
    $ouListBox.ValueMember = "DistinguishedName"
    $ouListBox.Items.AddRange($displayItems)
    $ouListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ouListBox.Font = New-Object System.Drawing.Font("΢���ź�", 9)
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
            
            $script:allUsersOU = $null  # ����������

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

#�½�OU��֯
function CreateNewOU {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("δѡ��ǰOU�������л���Ŀ��OU���ٲ���", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    #��ȡ�����Ĭ��Users����
    $domainRootDN = Get-DomainRootDN
    if (-not $domainRootDN) { return }
    $defaultUsersOU = "CN=Users,$domainRootDN"

    # ������OU���ƣ��߼����䣩
    $newOUName = [Microsoft.VisualBasic.Interaction]::InputBox(
		"��������OU�����ƣ�ʾ����ITDepartment�����񲿣�`n`nע�⣺���ɰ��������ַ���/\=+:*#$@?!~`"<>|��", 
		"�½�OU��֯", 
		""
    )

    # ����У�飨�߼����䣩
	if ([string]::IsNullOrEmpty($newOUName)) {
		return
	}
	elseif ([string]::IsNullOrWhiteSpace($newOUName)) {
		[System.Windows.Forms.MessageBox]::Show("OU���Ʋ���Ϊ�ջ�������ո�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
		return
	}
		
    # �����ַ�У�飨�߼����䣩
    $invalidChars = '[\\/=+:*#$@?!~"<>|]'
    if ($newOUName -match $invalidChars) {
        $matchedChar = $matches[0]
        [System.Windows.Forms.MessageBox]::Show("OU���ư����Ƿ��ַ���`"$matchedChar`"`n��ɾ�������ԣ�", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # �жϸ�����
    $parentDN = $null
    $parentDisplay = $null

    # ��ǰOU�ǡ�Ĭ��Users������ʱ��������Ϊ���
    if ($script:currentOU -eq $defaultUsersOU) {
        $parentDN = $domainRootDN
        $domainDNS = $domainRootDN -replace 'DC=','.' -replace ',',''
        $parentDisplay = "��� ($domainDNS)"
    }
    # ���������ֱ���Ե�ǰOU��Ϊ������
    else {
        $parentDN = $script:currentOU
        
        # ���ɸ������Ѻ���ʾ���ƣ��߼����䣩
        $ouParts = @()
        $parentDN -split ',' | ForEach-Object {
            if ($_ -match '^OU=(.+)') { $ouParts += $matches[1] }
            elseif ($_ -eq "CN=Users") { $ouParts += "Users����" }
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

        # �����ӳ�ȷ��ADͬ��
        Start-Sleep -Milliseconds 500

        # �����ɹ�����
        $script:currentOU = $newOUFullDN
        $script:textOU.Text = $script:currentOU

        $script:connectionStatus = "OU�����ɹ���$newOUFullDN"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU�����ɹ���`n����·����$newOUFullDN", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # ǿ��ˢ��OU�б�
        LoadOUList | Out-Null

        # ˢ���û������б�
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
        # ������
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

#������OU��֯
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

    # ��ȡ�ܱ��������������Ĭ��Users��
    $domainRootDN = Get-DomainRootDN
    if (-not $domainRootDN) { return }
    $defaultUsersOU = "CN=Users,$domainRootDN"
    $protectedContainers = @(
        $defaultUsersOU,  # ��̬Ĭ��Users����
        $domainRootDN     # ��̬�������
    )

    # ��鵱ǰOU�Ƿ�Ϊϵͳ�ؼ�����
    if ($protectedContainers -contains $script:currentOU) {
        $containerName = if ($script:currentOU -eq $domainRootDN) {
            $domainDNS = $domainRootDN -replace 'DC=','.' -replace ',',''
            "���������$domainDNS��"
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
            Rename-ADObject -Identity $oldDN -NewName $newName -ErrorAction Stop
        } -ArgumentList $script:currentOU, $newOUName -ErrorAction Stop

        # �������ɹ�����
        $script:currentOU = $newOUFullDN
        $script:textOU.Text = $script:currentOU
        
        $script:connectionStatus = "OU�������ɹ���$currentOUName -> $newOUName"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU�������ɹ���`n������·����$newOUFullDN", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
		Start-Sleep -Milliseconds 500
        # ˢ������б�
        LoadOUList
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
        # ������
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

#ɾ��OU��֯
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

    # ��ȡ�ܱ��������������Ĭ��Users��
    $domainRootDN = Get-DomainRootDN
    if (-not $domainRootDN) { return }
    $defaultUsersOU = "CN=Users,$domainRootDN"
    $protectedContainers = @(
        $defaultUsersOU,  # ��̬Ĭ��Users����
        $domainRootDN     # ��̬�������
    )

    # ��鵱ǰOU�Ƿ�Ϊϵͳ�ؼ�����
    if ($protectedContainers -contains $script:currentOU) {
        $containerName = if ($script:currentOU -eq $domainRootDN) {
            $domainDNS = $domainRootDN -replace 'DC=','.' -replace ',',''
            "���������$domainDNS��"
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

    # ��ʾOU�����Ϣ
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
                Write-Warning "�������ʱ����: $($_.Exception.Message)"
            }
            
            # �ݹ�ɾ��
            Remove-ADOrganizationalUnit -Identity $ouDN -Recursive -Confirm:$false -ErrorAction Stop
        } -ArgumentList $script:currentOU -ErrorAction Stop

        # �����ӳ�ȷ��ADͬ��
        Start-Sleep -Milliseconds 500

        # ɾ���ɹ�����
        $script:connectionStatus = "OUɾ���ɹ���$displayHierarchy"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU�ѳ���ɾ�����������Ӷ��󣩣�", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # ���õ�ǰOUΪ������
        $parentDN = $script:currentOU -replace '^[^,]+,', ''
        if ($parentDN -match '^OU=|^CN=Users,|^DC=') {  # ������������OU��Users���������
            $script:currentOU = $parentDN
        }
        else {
            $script:currentOU = $parentDN
        }
        $script:textOU.Text = $script:currentOU

        # ǿ��ˢ��OU�б�
        LoadOUList | Out-Null

        # ˢ���û������б�
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
        # ���������ʾ
        $errorMsg = $_.Exception.Message
        if ($script:currentOU -eq $defaultUsersOU) {
            $errorMsg = "�޷�ɾ��Ĭ��Users����������ϵͳ����������������ɾ����"
        } elseif ($script:currentOU -eq $domainRootDN) {
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
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

        # Զ�̻�ȡ����OU
        $script:allOUs = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            Import-Module ActiveDirectory -ErrorAction Stop
            Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName |
                Where-Object { $_.Name -ne "Domain Controllers" } |			
                Select-Object Name, DistinguishedName |
                Sort-Object Name
        } -ErrorAction Stop

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



# �л�OU��֯
function SwitchOU {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ����OU�б�
    $ous = LoadOUList
    if (-not $ous -or $ous.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("δ�ҵ��κ�OU��֯", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # ��ȡĬ��Users����������Users������Ϣ
    $defaultUsersOU = $null
    $allUsersOU = $null
    $domainDN = $null
    
    # ���Ի�ȡ����Ϣ
    if ($script:domainContext -and $script:domainContext.DomainInfo) {
        $domainDN = $script:domainContext.DomainInfo.DefaultPartition
    }
    # �ӵ�ǰOU��������Ϣ
    if (-not $domainDN -and $script:currentOU) {
        $domainDN = $script:currentOU -replace '^[^,]+,', ''
    }
    
    # ����Ĭ��Users����DN������Users����DN
    if ($domainDN) {
        $defaultUsersOU = "CN=Users,$domainDN"
        $script:allUsersOU = "CN=Users,$domainDN"
    }

    # �����̶�ѡ�������ӵ��б�
    $fixedItems = @()
    if (-not [string]::IsNullOrWhiteSpace($defaultUsersOU)) {
        $fixedItems += [PSCustomObject]@{
            Name = "Ĭ��Users"
            DistinguishedName = $defaultUsersOU
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($script:allUsersOU)) {
        $fixedItems += [PSCustomObject]@{
            Name = "����Users"
            DistinguishedName = $script:allUsersOU
        }
    }

    # �ϲ��̶�ѡ���ԭ��OU�б��̶�ѡ�����ǰ�棩
    $displayItems = $fixedItems + $ous

    # ����OUѡ��Ի���
    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "ѡ��OU��֯"
    $ouForm.Size = New-Object System.Drawing.Size(350, 250)  # �̶����ڴ�С
    $ouForm.StartPosition = "CenterScreen"
    $ouForm.MaximizeBox = $false
    $ouForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

    # ���������ð�ť���������У�
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = "Bottom"
    $buttonPanel.Height = 40
    $buttonPanel.Padding = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
    $buttonPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

    $ouListBox = New-Object System.Windows.Forms.ListBox
    $ouListBox.Dock = "Fill"
    $ouListBox.DisplayMember = "Name"
    $ouListBox.ValueMember = "DistinguishedName"
    $ouListBox.Items.AddRange($displayItems)
    $ouListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
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
    $okButton.Location = New-Object System.Drawing.Point(55, 5)
    $okButton.Add_Click({
        if ($ouListBox.SelectedItem) {
			$selectedItem = $ouListBox.SelectedItem
			$script:currentOU = $selectedItem.DistinguishedName
			$script:textOU.Text = $script:currentOU
			
			# ����ѡ������allUsersOU״̬
			if ($selectedItem.Name -eq "����Users") {
				# ѡ��"����Users"ʱ������allUsersOUΪ��ӦDN
				$script:allUsersOU = $selectedItem.DistinguishedName
			} else {
				# ����ѡ��ʱ�����allUsersOU����ʾֻ����ָ��OU��
				$script:allUsersOU = $null
			}
			
			$ouForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
			}
    })

    # ȡ����ť
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "ȡ��"
    $cancelButton.Width = 100
    $cancelButton.Height = 30
	$cancelButton.FlatAppearance.BorderSize = 1
    $cancelButton.Location = New-Object System.Drawing.Point(180, 5)
    $cancelButton.Add_Click({
        $ouForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    })

    # ��ӿؼ������ͱ�
    $buttonPanel.Controls.Add($okButton)
    $buttonPanel.Controls.Add($cancelButton)
    $ouForm.Controls.Add($ouListBox)
    $ouForm.Controls.Add($buttonPanel)

    if ($ouForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $displayName = switch -Wildcard ($script:currentOU) {
            "^CN=Users," { "Ĭ��Users����" }
            "^CN=Users," { "����Users����" }
            default { $script:currentOU.Split(',')[0] }
        }
        $script:connectionStatus = "���л���OU��$displayName"
        UpdateStatusBar
        
        # �л�OU��ˢ���û������б�
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


    $newOUName = [Microsoft.VisualBasic.Interaction]::InputBox(
		"��������OU�����ƣ�ʾ����ITDepartment��Finance��`n`nע�⣺���ɰ��������ַ���/ \ : * ? `" < > |��", 
		"�½�OU��֯", 
		""
    )


    # ������ֵУ��
	if ([string]::IsNullOrEmpty($newOUName)) {
		return
	}
	elseif ([string]::IsNullOrWhiteSpace($newOUName)) {
		[System.Windows.Forms.MessageBox]::Show("OU���Ʋ���Ϊ�ջ�������ո�", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
		return
	}


    # �����ַ�����У��
    $invalidChars = '[\\/:*?"<>|]'
    if ($newOUName -match $invalidChars) {
        $matchedChar = $matches[0]
        [System.Windows.Forms.MessageBox]::Show("OU���ư����Ƿ��ַ���`"$matchedChar`"`n��ɾ�������ԣ�", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # �Զ���ȡ���Ĭ�Ϸ���
    $domainDN = $null
    $logMessages = @("��ʼ��ȡ����Ϣ...")
    
    # ����1: ��domainContext��ȡ����ϸ��ϣ�
    if (-not $domainDN) {
        try {
            $logMessages += "���Դ�domainContext��ȡ����Ϣ..."
            if ($script:domainContext) {
                $logMessages += "domainContext����"
                if ($script:domainContext.DomainInfo) {
                    $logMessages += "DomainInfo���Դ���"
                    $domainDN = $script:domainContext.DomainInfo.DefaultPartition
                    $logMessages += "��domainContext��ȡ��: $domainDN"
                }
                else {
                    $logMessages += "domainContext�в�����DomainInfo����"
                }
            }
            else {
                $logMessages += "script:domainContextΪ��"
            }
        }
        catch {
            $logMessages += "��domainContext��ȡʧ��: $($_.Exception.Message)"
        }
    }

    # ����2: ��Զ�̻Ự��ȡ��ֱ�ӹ���Ԥ�ڸ�ʽ��
    if (-not $domainDN -and $script:remoteSession) {
        try {
            $logMessages += "���Դ�Զ�̻Ự��ȡ����Ϣ..."
            # ����2.1: ʹ��Get-ADDomain
            $domainInfo = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                Import-Module ActiveDirectory -ErrorAction Stop
                Get-ADDomain -ErrorAction Stop
            } -ErrorAction Stop
            
            if ($domainInfo -and $domainInfo.DefaultPartition) {
                $domainDN = $domainInfo.DefaultPartition
                $logMessages += "��Get-ADDomain��ȡ��: $domainDN"
            }
            else {
                $logMessages += "Get-ADDomainδ������Ч��Ϣ"
                # ����2.2: ���Ի�ȡ��ǰ���������������ת��ΪDN
                $domainController = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                    $env:USERDNSDOMAIN
                } -ErrorAction Stop
                
                if ($domainController) {
                    $logMessages += "��ȡ���������: $domainController"
                    $domainParts = $domainController -split '\.'
                    if ($domainParts.Count -ge 2) {
                        $domainDN = "dc=$($domainParts -join ',dc=')"
                        $logMessages += "ת����õ���DN: $domainDN"
                    }
                }
            }
        }
        catch {
            $logMessages += "��Զ�̻Ự��ȡʧ��: $($_.Exception.Message)"
        }
    }

    # ����3: �ӵ�ǰOU����������е�ǰOU��
    if (-not $domainDN -and -not [string]::IsNullOrWhiteSpace($script:currentOU)) {
        try {
            $logMessages += "���Դӵ�ǰOU��������Ϣ: $script:currentOU"
            # �ӵ�ǰOU����ȡ����Ϣ
            $domainDN = $script:currentOU -replace '^[^,]+,', ''
            # ��֤�Ƿ�����Ч�����ʽ
            if ($domainDN -match '^DC=') {
                $logMessages += "�ӵ�ǰOU�����õ�: $domainDN"
            }
            else {
                $logMessages += "���������ʽ��Ч: $domainDN"
                $domainDN = $null
            }
        }
        catch {
            $logMessages += "�ӵ�ǰOU����ʧ��: $($_.Exception.Message)"
        }
    }

    # ����4: ���ԴӼ������������ȡ
    if (-not $domainDN) {
        try {
            $logMessages += "���Դӱ��ؼ������������ȡ..."
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $domainName = $domain.Name
            $logMessages += "��ǰ������������: $domainName"
            
            # ת��ΪDN��ʽ
            $domainParts = $domainName -split '\.'
            if ($domainParts.Count -ge 2) {
                $domainDN = "dc=$($domainParts -join ',dc=')"
                $logMessages += "ת����õ���DN: $domainDN"
            }
        }
        catch {
            $logMessages += "�ӱ��ؼ������ȡ����Ϣʧ��: $($_.Exception.Message)"
        }
    }

    # ��������Զ�������ʧ�ܣ���¼��־����ʾ�ֶ�����
    if (-not $domainDN) {
        # ����־���浽��ʱ�ļ��Ա��Ų�����
        $logPath = "$env:TEMP\DomainInfoLog.txt"
        $logMessages | Out-File -FilePath $logPath -Encoding utf8
        
        $domainDN = [Microsoft.VisualBasic.Interaction]::InputBox(
            "�޷��Զ���ȡ����Ϣ����־�ѱ��浽: $logPath��`n���ֶ��������DistinguishedName`n���磺DC=abc-test,DC=com", 
            "�ֶ���������Ϣ", 
            "dc=abc-test,dc=com"  # Ԥ���û�����
        )
        
        if ([string]::IsNullOrWhiteSpace($domainDN) -or $domainDN -notmatch '^DC=') {
            [System.Windows.Forms.MessageBox]::Show("��Ч������Ϣ����ʽӦΪDC=domain,DC=com", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }

    # ����У��Path����
    if ([string]::IsNullOrWhiteSpace($domainDN)) {
        [System.Windows.Forms.MessageBox]::Show("�޷���ȡ��Ч������Ϣ���޷�����OU��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ȷ����OU�ĸ�����
    $newOUFullDN = "OU=$newOUName,$domainDN"
    $confirmMsg = "ȷ��������λ�ô���OU��`n��������$domainDN`n��OU����·����$newOUFullDN"
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
        } -ArgumentList $newOUName, $domainDN -ErrorAction Stop

        # �����ɹ�����
        $script:connectionStatus = "OU�����ɹ���$newOUFullDN"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU�����ɹ���`n����·����$newOUFullDN", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # �л����´�����OU
        $script:currentOU = $newOUFullDN
        $script:textOU.Text = $script:currentOU

        # ˢ���û������б�
        LoadUserList
        LoadGroupList

    } catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "already exists") {
            $errorMsg = "OU�Ѵ��ڣ���������ƣ��磺$newOUName_2��"
        } elseif ($errorMsg -match "permission") {
            $errorMsg = "Ȩ�޲��㣡��ȷ�Ϲ���Ա�˺�ӵ��OU����Ȩ��"
        } elseif ($errorMsg -match "Path") {
            $errorMsg = "·����������$domainDN����������Ϣ�Ƿ���ȷ"
        }
        $script:connectionStatus = "OU����ʧ�ܣ�$errorMsg"
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

    # У�鵱ǰ�Ƿ���ѡ�е�OU
    if ([string]::IsNullOrWhiteSpace($script:currentOU)) {
        [System.Windows.Forms.MessageBox]::Show("δ��ȡ����ǰOU��Ϣ��������������أ�", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # ����ȷ�ϣ���Σ������
    $confirmMsg = "���棺ɾ��OU��ͬʱɾ���������ж����û����顢��OU����`nȷ��Ҫɾ������OU��`n$script:currentOU"
    $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "��Σ����ȷ��", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $script:connectionStatus = "����ɾ��OU��$script:currentOU..."
        UpdateStatusBar
        $script:mainForm.Refresh()

        # Զ��ɾ��OU
        Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($ouDN)
            Import-Module ActiveDirectory -ErrorAction Stop
            
            # ������������Ӵ�����
            try {
                Set-ADOrganizationalUnit -Identity $ouDN -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
            }
            catch {
                Write-Warning "�������ʱ����: $($_.Exception.Message)"
            }
            
            # �ݹ�ɾ��OU�������Ӷ���
            Remove-ADOrganizationalUnit -Identity $ouDN -Recursive -Confirm:$false -ErrorAction Stop
        } -ArgumentList $script:currentOU -ErrorAction Stop

        # ɾ���ɹ�����
        $script:connectionStatus = "OUɾ���ɹ���$script:currentOU"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show("OU�ѳ���ɾ�������Ӷ��󣩣�", "�ɹ�", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # ���õ�ǰOUΪ���Ĭ��Users OU��ʹ�ø��ɿ�������Ϣ��ȡ��ʽ��
        $domainDN = $null
        # ���Դ�domainContext��ȡ����Ϣ
        if ($script:domainContext -and $script:domainContext.DomainInfo) {
            $domainDN = $script:domainContext.DomainInfo.DefaultPartition
        }
        # ���ʧ�ܣ����Դӵ�ǰOU����
        if (-not $domainDN -and $script:currentOU) {
            $domainDN = $script:currentOU -replace '^[^,]+,', ''
        }
        
        # �����ȡ������Ϣ������Ĭ��Users OU
        if ($domainDN) {
            $defaultUsersOU = "CN=Users,$domainDN"
            $script:currentOU = $defaultUsersOU
            $script:textOU.Text = $script:currentOU
        }
        else {
            $script:currentOU = $null
            $script:textOU.Text = ""
        }

        # ˢ���û������б�
        LoadUserList
        LoadGroupList

    } catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "not found") {
            $errorMsg = "OU�����ڣ������ѱ�ɾ����·������"
        } elseif ($errorMsg -match "permission") {
            $errorMsg = "Ȩ�޲��㣡��ȷ�Ϲ���Ա�˺�ӵ��OUɾ��Ȩ��"
        }
		Write-Error $errorMsg
        $script:connectionStatus = "OUɾ��ʧ�ܣ�$errorMsg"
        UpdateStatusBar
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}


    
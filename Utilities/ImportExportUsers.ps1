<#
.SYNOPSIS
AD�û��������������
#>


function ImportCSVAndCreateUsers {	
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }	

    # ���ļ�ѡ��Ի���
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "CSV�ļ� (*.csv)|*.csv|�����ļ� (*.*)|*.*"
    $fileDialog.Title = "ѡ������û���Ϣ��CSV�ļ�"
    
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvPath = $fileDialog.FileName
        $script:connectionStatus = "���ڴ���CSV�ļ�: $([System.IO.Path]::GetFileName($csvPath))"
        $script:mainForm.Refresh()

        try {
            # ���CSV�ļ��Ƿ����
            if (-not (Test-Path -Path $csvPath -PathType Leaf)) {
                [System.Windows.Forms.MessageBox]::Show("CSV�ļ�������: $csvPath", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            # ��ȡCSV�ļ�
            try {
                $users = Import-Csv -Path $csvPath -Encoding Default -ErrorAction Stop
                $userCount = $users | Measure-Object | Select-Object -ExpandProperty Count
                $script:connectionStatus = "�ɹ���ȡCSV�ļ��������� $userCount ���û���¼"
                $script:mainForm.Refresh()
            } catch {
                [System.Windows.Forms.MessageBox]::Show("��ȡCSV�ļ�ʧ��: $_", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            # ��ʾȷ�϶Ի���
            $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                "���� $userCount ���˺ż�¼���Ƿ�ȷ�����벢������Щ�û���",
                "ȷ�ϵ���",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
                $script:connectionStatus = "�û�ȡ���˵������"
                return
            }

            # ���û��������л�Ϊ�ַ������Ա���Զ�̻Ự��ʹ��
            $usersJson = $users | ConvertTo-Json
            
            $script:connectionStatus = "����Զ�̴����û�..."
            $script:mainForm.Refresh()

            # ִ��Զ�̲���
            $result = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($usersJson, $NameOU)
                
                # ��ʼ��������������ϸ��Ϣ����
                $result = [PSCustomObject]@{
                    TotalUsers = 0
                    CreatedUsers = 0
                    SkippedUsers = 0
                    ExistingUsers = @()  # �洢�Ѵ��ڵ��û���Ϣ
                    TotalGroups = 0
                    CreatedGroups = 0
                    CreatedGroupsDetails = @()  # �洢�����ɹ�������Ϣ
                    ErrorLogs = @()
                    CreatedUsersDetails = @()  # �洢�����ɹ����û�����
                }

                # �����л��û�����
                try {
                    $users = $usersJson | ConvertFrom-Json
                    $result.TotalUsers = $users.Count
                }
                catch {
                    $result.ErrorLogs += "�����û�����ʧ��: $_"
                    return $result
                }

                # ���ADģ��
                if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
                    $result.ErrorLogs += "Active Directoryģ��δ��װ���밲װRSAT�����е�ADģ�顣"
                    return $result
                }
                try {
                    Import-Module ActiveDirectory -ErrorAction Stop
                }
                catch {
                    $result.ErrorLogs += "����Active Directoryģ��ʧ��: $_"
                    return $result
                }

                # -------------------------- �����б�����ԭ���ܣ� --------------------------
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
                # -------------------------------------------------------------------

                # ����û������Ƿ���Group����
                $hasGroupProperty = $false
                if ($users.Count -gt 0) {
                    $firstUser = $users[0]
                    $hasGroupProperty = $firstUser.PSObject.Properties.Name -contains "Group"
                }

                # �ռ��������������
                if ($hasGroupProperty) {
                    $groups = $users | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Group) } | Select-Object -ExpandProperty Group -Unique
                    $result.TotalGroups = $groups.Count

                    if ($groups.Count -gt 0) {
                        foreach ($group in $groups) {
                            try {
                                # ������Ƿ����
                                $existingGroup = Get-ADGroup -Identity $group -ErrorAction Stop
                            }
                            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                                # �鲻���ڣ����Դ���
                                try {
                                    if ([string]::IsNullOrWhiteSpace($NameOU)) {
                                        $result.ErrorLogs += "�޷�ȷ���� $group ��OU����������"
                                        continue
                                    }
                                    
                                    # ��֤OU�Ƿ����
                                    if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$NameOU'" -ErrorAction SilentlyContinue) -and
                                        -not (Get-ADObject -Filter "DistinguishedName -eq '$NameOU'" -ErrorAction SilentlyContinue)) {
                                        $result.ErrorLogs += "���Ŀ��OU������: $NameOU���޷������� $group"
                                        continue
                                    }
                                    
                                    # ���������
                                    $groupParams = @{
                                        Name            = $group
                                        SamAccountName  = $group
                                        GroupCategory   = "Security"
                                        GroupScope      = "Global"
                                        Path            = $NameOU
                                        Description     = "��������������Զ�������: $group"
                                        ErrorAction     = "Stop"
                                    }
                                    
                                    New-ADGroup @groupParams
                                    $result.CreatedGroups++
                                    $result.CreatedGroupsDetails += $group  # ��¼�����ɹ�������
                                }
                                catch {
                                    $result.ErrorLogs += "������ $group ʧ��: $($_.Exception.Message)"
                                }
                            }
                            catch {
                                $result.ErrorLogs += "����� $group ʱ��������: $_"
                            }
                        }
                    }
                }

                # �����û�
                foreach ($user in $users) {
                    try {
                        # ����û��Ƿ��Ѵ���
                        $existingUser = Get-ADUser -Identity $user.SamAccountName -Properties DisplayName -ErrorAction Stop
                        
                        # ��¼�Ѵ��ڵ��û���Ϣ
                        $result.ExistingUsers += [PSCustomObject]@{
                            SamAccountName = $existingUser.SamAccountName
                            DisplayName = $existingUser.DisplayName
                        }
                        
                        $result.ErrorLogs += "�û� $($user.SamAccountName) �Ѵ��ڣ���������"
                        $result.SkippedUsers++
                        continue
                    }
                    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                        # �û������ڣ�������������
                    }
                    catch {
                        $result.ErrorLogs += "����û� $($user.SamAccountName) ʱ��������: $_"
                        $result.SkippedUsers++
                        continue
                    }
                    
                    # �ֶ���֤
                    $requiredFields = 'SamAccountName'
                    if ($missingFields = $requiredFields | Where-Object { [string]::IsNullOrWhiteSpace($user.$_) }) {
                        $result.ErrorLogs += "�û���¼ȱ�ٱ����ֶ�: $($missingFields -join ', ')������"
                        $result.SkippedUsers++
                        continue
                    }

                    # ��������
                    $plainPassword = if (-not [string]::IsNullOrWhiteSpace($user.Password)) { 
                        $user.Password 
                    } else { 
                        "P@ssw0rd$(Get-Random -Minimum 1000 -Maximum 9999)" 
                    }
                    $securePassword = ConvertTo-SecureString $plainPassword -AsPlainText -Force

                    #�˺Ź������ڴ���
                    $accountExpirationDate = $null
                    if (-not [string]::IsNullOrWhiteSpace($user.AccountExpirationDate)) {
                        try {
                            $date = [DateTime]::MinValue
                            # ���ȳ���Ĭ�ϸ�ʽ����
                            $parsed = [DateTime]::TryParse($user.AccountExpirationDate, [ref]$date)
                            
                            if (-not $parsed) {
                                # ���Զ��ֳ������ڸ�ʽ����
                                $formats = @(
                                    "yyyy.MM.dd", "MM/dd/yyyy", "dd/MM/yyyy",
                                    "yyyy-MM-dd", "yyyyMMdd", "yyyy/MM/dd",
                                    "yyyy��MM��dd��"
                                )
                                foreach ($format in $formats) {
                                    if ([DateTime]::TryParseExact(
                                        $user.AccountExpirationDate, 
                                        $format, 
                                        [System.Globalization.CultureInfo]::InvariantCulture, 
                                        [System.Globalization.DateTimeStyles]::None, 
                                        [ref]$date
                                    )) {
                                        $parsed = $true
                                        break
                                    }
                                }
                            }
                            
                            if ($parsed) {
                                # AD�衰��������+1�족�������û�����2024-10-05��AD��Ϊ2024-10-06��10-06��ҹ���ڣ�
                                $accountExpirationDate = $date.Date.AddDays(1)
                            } else {
                                $result.ErrorLogs += "�޷�������������: $($user.AccountExpirationDate)���������ù�������"
                            }
                        }
                        catch {
                            $result.ErrorLogs += "������������ʧ��: $($user.AccountExpirationDate)���������ù�������"
                        }
                    }

                    #���ղ���߼�
                    $fullName = $null
                    if (-not [string]::IsNullOrWhiteSpace($user.Surname) -and -not [string]::IsNullOrWhiteSpace($user.GivenName)) {
                        $surname = $user.Surname.Trim()
                        $givenName = $user.GivenName.Trim()
                    } else {
                        # ��DisplayName��Name�ֶλ�ȡ����������CSV��������һ����
                        if (-not [string]::IsNullOrWhiteSpace($user.DisplayName)) {
                            $fullName = $user.DisplayName.Trim()
                        } elseif (-not [string]::IsNullOrWhiteSpace($user.Name)) {
                            $fullName = $user.Name.Trim()
                        } else {
                            $result.ErrorLogs += "�û� $($user.SamAccountName) ȱ�������ֶΣ�DisplayName/Name�����޷�����պ���"
                            $result.SkippedUsers++
                            continue
                        }

                        # ����ʶ������
                        if ($fullName.Length -ge 2 -and $doubleSurnames -contains $fullName.Substring(0, 2)) {
                            $surname = $fullName.Substring(0, 2)
                            $givenName = if ($fullName.Length -gt 2) { $fullName.Substring(2).Trim() } else { "" }
                        } else {
                            $surname = $fullName.Substring(0, 1)
                            $givenName = if ($fullName.Length -gt 1) { $fullName.Substring(1).Trim() } else { $fullName }
                        }
                    }
                    # -------------------------------------------------------------------

                    # ��ȡ��ǰ����Ϣ
                    $domain = Get-ADDomain
                    
                    # �û�����
					$userParams = @{
						SamAccountName        = $user.SamAccountName
						UserPrincipalName     = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { "$($user.SamAccountName)@$($domain.DNSRoot)" }
						Name                  = $user.SamAccountName
						DisplayName           = if ($user.DisplayName) { $user.DisplayName } else { if ($fullName) { $fullName } else { $user.SamAccountName } }
						Surname               = $surname
						GivenName             = $givenName
						Path                  = $NameOU
						AccountPassword       = $securePassword
						Enabled               = $true
						ChangePasswordAtLogon = $true
						Description           = $user.Description
						EmailAddress          = $user.EmailAddress
						OfficePhone           = if (-not [string]::IsNullOrWhiteSpace($user.Telephone)) { $user.Telephone } else { $user.Phone }
						ErrorAction           = "Stop"
					}

                    # ֻ�н����ɹ�ʱ����ӹ������ڲ���
                    if ($accountExpirationDate) {
                        $userParams['AccountExpirationDate'] = $accountExpirationDate
                    }

                    # ����ղ���
                    $keysToRemove = $userParams.Keys | Where-Object { 
                        $null -eq $userParams[$_] -or [string]::IsNullOrWhiteSpace($userParams[$_]) 
                    }
                    $keysToRemove | ForEach-Object { $userParams.Remove($_) }

                    # �����û�
                    try {
                        New-ADUser @userParams
                        $result.CreatedUsers++
                        
                        # ��¼�����ɹ����û�����
                        $result.CreatedUsersDetails += [PSCustomObject]@{
                            Index = $result.CreatedUsers  # ��¼���
                            SamAccountName = $user.SamAccountName
                            DisplayName = $userParams.DisplayName
                            Surname = $surname
                            GivenName = $givenName
							Email = $user.EmailAddress
							Telephone = if (-not [string]::IsNullOrWhiteSpace($user.Telephone)) { $user.Telephone } else { $user.Phone }
                            Password = $plainPassword
                        }
                        
                        # ���û���ӵ���
                        if ($hasGroupProperty -and -not [string]::IsNullOrWhiteSpace($user.Group)) {
                            if (Get-ADGroup -Identity $user.Group -ErrorAction SilentlyContinue) {
                                Add-ADGroupMember -Identity $user.Group -Members $user.SamAccountName -ErrorAction Stop
                            } else {
                                $result.ErrorLogs += "�� $($user.Group) �����ڣ��޷����û� $($user.SamAccountName) ��ӵ�����"
                            }
                        }
                    } catch {
                        $result.ErrorLogs += "�����û� $($user.SamAccountName) ʧ��: $($_.Exception.Message)"
                        $result.SkippedUsers++
                    }
                }
                
                return $result
            } -ArgumentList $usersJson, $script:currentOU -ErrorAction Stop

            # ƴ�ӽ����Ϣ
            $msg = @"
����������ɣ�
========================================
���û�����$($result.TotalUsers)
�ɹ�������$($result.CreatedUsers)
�����û���$($result.SkippedUsers)
����������$($result.CreatedGroups)/$($result.TotalGroups)
========================================

"@

            # ��ӳɹ���������
            if ($result.CreatedGroups -gt 0) {
                $msg += "�ɹ��������飺`n$($result.CreatedGroupsDetails -join '��')`n`n"
				LoadGroupList
            }

            # ����Ѵ��ڵ��û�
            if ($result.ExistingUsers.Count -gt 0) {
                $msg += "�Ѵ��ڵ��û���`n"
                $result.ExistingUsers | ForEach-Object {
                    $msg += "$($_.SamAccountName)��$($_.DisplayName)��`n"
                }
                $msg += "`n"
            }

            # ��Ӵ����ɹ����û�
            if ($result.CreatedUsers -gt 0) {
                $msg += "�����ɹ����û�����ʼ���룩��`n"
                $result.CreatedUsersDetails | ForEach-Object {
                    $msg += "$($_.SamAccountName) | ������$($_.DisplayName) | ���룺$($_.Password)`n"
                }
                $msg += "`n"
            }

            # ��Ӵ�����Ϣ
			if ($result.ErrorLogs.Count -gt 0) {
				$msg += "`n������Ϣ��`n$($result.ErrorLogs -join "`r`n")`n"
			}

            # ��ʾ����Ի���
            [System.Windows.Forms.MessageBox]::Show($msg, "�����������", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

            # ˢ���û��б�
            LoadUserList
            $script:connectionStatus = "���������û����������"
        }
        catch {
            $script:connectionStatus = "���������û�ʧ��: $($_.Exception.Message)"
            $statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
            [System.Windows.Forms.MessageBox]::Show("ִ����������ʱ��������: $($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}


function ExportCSVUsers {	
    # 1. ǰ��У�飺������+Զ�̻Ự�Ƿ���Ч
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    if (-not $script:remoteSession -or $script:remoteSession.State -ne "Opened") {
        [System.Windows.Forms.MessageBox]::Show("Զ�̻Ựδ�������ѶϿ����������������", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 2. �����ļ�����Ի��򣨿ͻ���ѡ��·����
    $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $fileDialog.Filter = "CSV�ļ� (*.csv)|*.csv|�����ļ� (*.*)|*.*"
    $fileDialog.Title = "ѡ��AD�û���Ϣ�ĵ���·��"
    $fileDialog.DefaultExt = "csv"
    $fileDialog.AddExtension = $true
    # Ĭ�ϵ���������
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($fileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        $script:connectionStatus = "�û�ȡ���˵�������"
        return
    }

    $csvPath = $fileDialog.FileName
    $script:connectionStatus = "���ڴ���ض�ȡ�û�����..."
    $script:mainForm.Refresh()

    try {
        # 3. Զ�̻Ự������ȡAD���ݣ����漰�ļ�������
        $remoteUserData = Invoke-Command -Session $script:remoteSession -ScriptBlock {
			param($NameOU, $allUsersOU)
            # ��ʼ��Զ�̽������
            $remoteResult = [PSCustomObject]@{
                UserData = $null
                ErrorLogs = @()
            }

            # ���ADģ��
            if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
                $remoteResult.ErrorLogs += "���δ��װActive Directoryģ�飨�谲װRSAT���ߣ�"
                return $remoteResult
            }
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
            }
            catch {
                $remoteResult.ErrorLogs += "����ADģ��ʧ�ܣ�$($_.Exception.Message)"
                return $remoteResult
            }

            # ��ȡAD�û�������Surname��GivenName���ԣ�
            $adProperties = @(
                "SamAccountName", "DisplayName", "UserPrincipalName",
                "EmailAddress", "TelephoneNumber", "Description", "GivenName", "Surname",
                "AccountExpirationDate", "Enabled", "DistinguishedName",
                "Department", "Title", "OfficePhone", "LastLogonDate"
            )
            try {
				
				if ($allUsersOU) {
					$users = Get-ADUser -Filter * -Properties $adProperties -ErrorAction Stop
				} else {
					# ɸѡ�����������ض�OU
					$users = Get-ADUser -Filter * -SearchBase $NameOU -Properties $adProperties -ErrorAction Stop
				}
                $remoteResult.UserData = $users | Select-Object $adProperties  # ��������Ҫ������
            }
            catch {
                $remoteResult.ErrorLogs += "��ȡAD�û�ʧ�ܣ�$($_.Exception.Message)"
                return $remoteResult
            }

            return $remoteResult
        } -ArgumentList $script:currentOU, $script:allUsersOU -ErrorAction Stop

        # 4. ���Զ�̶�ȡ�Ƿ�ɹ�
        if ($remoteUserData.ErrorLogs.Count -gt 0) {
            throw "Զ�̶�ȡ����ʧ�ܣ�`n$($remoteUserData.ErrorLogs -join "`n")"
        }
        if (-not $remoteUserData.UserData -or $remoteUserData.UserData.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("δ����ض�ȡ���κ�AD�û�����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $script:connectionStatus = "����������ȡ�������û����ݣ�"
            return
        }

        # 5. ���ش�������
        $script:connectionStatus = "���ڱ�������CSV�ļ�..."
        $script:mainForm.Refresh()

        $exportData = $remoteUserData.UserData | ForEach-Object {
            [PSCustomObject]@{
                "�û���(SamAccountName)"    = $_.SamAccountName
                "��ʾ����"                  = $_.DisplayName
                "��(Surname)"               = $_.Surname
                "��(GivenName)"             = $_.GivenName
                "�����ַ"                  = $_.EmailAddress
				"�绰(TelephoneNumber)"     = if ($_.TelephoneNumber) { $_.TelephoneNumber } else { "��" }
                "����"                      = $_.Description
                "�˺Ź�������"              = if ($_.AccountExpirationDate) { $_.AccountExpirationDate.AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss") } else { "��" }
                "����¼ʱ��"              = if ($_.LastLogonDate) { $_.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "��δ��¼" }
                "�˺�״̬"                  = if ($_.Enabled) { "����" } else { "����" }
                "����OU"                    = $_.DistinguishedName
                "����"                      = $_.Department
                "ְλ"                      = $_.Title
                "�칫�绰"                  = $_.OfficePhone
            }
        }

        # 6. ���ص���CSV
        try {
            # У�鵼��·����д��Ȩ��
            $exportDir = [System.IO.Path]::GetDirectoryName($csvPath)
            if (-not (Test-Path -Path $exportDir -PathType Container)) {
                throw "����Ŀ¼�����ڣ�$exportDir"
            }
            # ����д��Ȩ��
            $testFile = [System.IO.Path]::Combine($exportDir, "test_permission.tmp")
            New-Item -Path $testFile -ItemType File -Force | Out-Null
            Remove-Item -Path $testFile -Force | Out-Null

            # ��ʽ������UTF8����������ģ�
            $exportData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
        }
        catch {
            throw "���ص���CSVʧ�ܣ�$($_.Exception.Message)"
        }

        # 7. �����ɹ�����ʾ+���ļ�ѡ��
        $msg = @"
���������ɹ���
========================================
�����ļ���$csvPath
�����û�����$($exportData.Count)
========================================
�Ƿ��������ļ��鿴��
"@
        $openResult = [System.Windows.Forms.MessageBox]::Show($msg, "�������", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
        if ($openResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            Start-Process -FilePath $csvPath  # ��Ĭ�ϳ����CSV����Excel��
        }

        $script:connectionStatus = "AD�û��������������"
    }
    catch {
        # ����ͳһ����
        $errorMsg = "����ʧ�ܣ�$($_.Exception.Message)"
        $script:connectionStatus = $errorMsg
        $statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

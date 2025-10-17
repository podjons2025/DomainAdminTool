<# 
���ĺ������������Ƶ�¼ʱ�䴰��
#>
function ShowRestrictLogonTimeForm {
    # ǰ�ü�飺�������״̬
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 1. ��ȡ���һ����ʼ�գ���̬���䲻ͬ������ã�
    try {
        $domainFirstDayOfWeek = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            (Get-WinSystemLocale).DateTimeFormat.FirstDayOfWeek.value__
        } -ErrorAction Stop

        # ��̬���ɡ�AD�ֽ����������ڡ�ӳ��
        $script:adWeekDays = switch ($domainFirstDayOfWeek) {
            0 { @("����", "��һ", "�ܶ�", "����", "����", "����", "����") }
            1 { @("��һ", "�ܶ�", "����", "����", "����", "����", "����") }
            default { @("��һ", "�ܶ�", "����", "����", "����", "����", "����") }
        }

        # ��̬���ɡ�UIѡ�����ڡ�AD������ӳ��
        $script:uiToAdWeekDay = @{}
        for ($i = 0; $i -lt $script:adWeekDays.Count; $i++) {
            $script:uiToAdWeekDay[$script:adWeekDays[$i]] = $i
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("��ȡ���ʱ������ʧ�ܣ�$($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # 2. ����û�ѡ��
    $selectedUsers = $script:userDataGridView.SelectedRows
    if (-not $selectedUsers -or $selectedUsers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "�������û��б���ѡ��1�������û�", 
            "��ʾ", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    try {
        $adminUsers = @()  # �洢��⵽�������Ա�û�
        # ����ÿ��ѡ���û�����ѯ���Ƿ�����Domain Admins��
        foreach ($row in $selectedUsers) {
            $samAccountName = $row.Cells["SamAccountName"].Value
            $displayName = $row.Cells["DisplayName"].Value
            $displayName = if ([string]::IsNullOrEmpty($displayName)) { $samAccountName } else { $displayName }

            # Զ�̲�ѯ�û����������а�ȫ�飨��Ƕ���飩
            $userGroups = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($sam)
                Import-Module ActiveDirectory -ErrorAction Stop
                # ��ȡ�û������а�ȫ�飬ɸѡ��������"Domain Admins"����
                Get-ADPrincipalGroupMembership -Identity $sam -ErrorAction Stop | 
                    Where-Object { $_.GroupCategory -eq "Security" -and $_.Name -eq "Domain Admins" } | 
                    Select-Object -ExpandProperty Name
            } -ArgumentList $samAccountName -ErrorAction Stop

            # ������Domain Admins�飬��¼���û�
            if ($userGroups -contains "Domain Admins") {
                $adminUsers += "$displayName���˺ţ�$samAccountName��"
            }
        }

        # ����⵽�����Ա�������澯����ֹ����
        if ($adminUsers.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "���棺ѡ�е��û�����������ع���Ա����ֹ�������¼Ȩ�ޣ�`n`n�漰�û���`n$($adminUsers -join "`n")`n`n������ѡ��ǹ���Ա�û�������", 
                "��Σ��������", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error  # �ô���ͼ��ǿ����ʾ
            )
            return  # ��ֹ��������������������
        }
    }
    catch {
        # �����ѯADʱ���쳣����Ȩ�޲��㡢�������⣩
        [System.Windows.Forms.MessageBox]::Show(
            "����Ա���У��ʧ�ܣ�$($_.Exception.Message)", 
            "����", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # 3. ������������
    $restrictTimeForm = New-Object System.Windows.Forms.Form
    $restrictTimeForm.Text = "�����û���¼ʱ��(֧���˺Ŷ�ѡ)"
    $restrictTimeForm.Size = New-Object System.Drawing.Size(700, 700)
    $restrictTimeForm.StartPosition = "CenterParent"
    $restrictTimeForm.FormBorderStyle = "FixedDialog"
    $restrictTimeForm.MaximizeBox = $false
    $restrictTimeForm.MinimizeBox = $false

    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 5000
    $toolTip.InitialDelay = 1000
    $toolTip.ReshowDelay = 500
    $toolTip.ShowAlways = $true

    # 4. ���ڲ��֣�4�У����⸴ѡ���������ȫѡ���ʱ����б���ײ���ť��
    $mainTable = New-Object System.Windows.Forms.TableLayoutPanel
    $mainTable.Dock = "Fill"
    $mainTable.Padding = 10
    $mainTable.ColumnCount = 1
    $mainTable.RowCount = 4
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50)))    # ��������ʱ�临ѡ��
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))    # ������ȫѡ��ѡ��
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))   # ʱ����б�
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)))    # �ײ���ť

    # 5. ����������ʱ���¼����ѡ��
    $chkAllowAllTime = New-Object System.Windows.Forms.CheckBox
    $chkAllowAllTime.Text = "��������ʱ���¼�������·�ʱ��ѡ��"
    $chkAllowAllTime.Font = New-Object System.Drawing.Font($chkAllowAllTime.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $chkAllowAllTime.AutoSize = $false
    $chkAllowAllTime.Width = $restrictTimeForm.ClientSize.Width - 40
    $chkAllowAllTime.Height = 40
    $chkAllowAllTime.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $chkAllowAllTime.Padding = New-Object System.Windows.Forms.Padding(2, 5, 2, 5)
    
    $chkAllowAllTime.Add_CheckedChanged({
        $lstLogonTime.Enabled = -not $chkAllowAllTime.Checked
        # ͬ������/���ù�����ȫѡ��
        $script:weekDayCheckboxes | ForEach-Object { $_.Enabled = -not $chkAllowAllTime.Checked }
    })
    $mainTable.Controls.Add($chkAllowAllTime, 0, 0)

    # 6. ������ȫѡ��ѡ���޸���ColumnStyle�İٷֱȼ������⣩
    $weekDayTable = New-Object System.Windows.Forms.TableLayoutPanel
    $weekDayTable.Dock = "Fill"
    $weekDayTable.ColumnCount = 7  # 7�������ո�ռһ��
    
    # ���Ĭ������ʽ
    while ($weekDayTable.ColumnStyles.Count -gt 0) {
        $weekDayTable.ColumnStyles.RemoveAt(0)
    }
    
    # ����ÿ�п�Ȱٷֱȣ�ʹ����ʽ����������������
    $columnPercent = 100.0 / 7  # ��ʽʹ�ø���������
    for ($i = 0; $i -lt 7; $i++) {
        $colStyle = New-Object System.Windows.Forms.ColumnStyle
        $colStyle.SizeType = [System.Windows.Forms.SizeType]::Percent
        $colStyle.Width = $columnPercent  # ʹ��Ԥ����İٷֱ�ֵ
        $weekDayTable.ColumnStyles.Add($colStyle)
    }

    # ����UI�̶���ʾ�Ĺ�����
    $script:weekDaysUI = @("��һ", "�ܶ�", "����", "����", "����", "����", "����")
    $script:weekDayCheckboxes = @()  # �洢���й����ո�ѡ��

    foreach ($day in $script:weekDaysUI) {
        $chkDay = New-Object System.Windows.Forms.CheckBox
        $chkDay.Text = $day
        $chkDay.Font = New-Object System.Drawing.Font($chkDay.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
        $chkDay.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $chkDay.AutoSize = $true
        $toolTip.SetToolTip($chkDay, "ȫѡ/ȡ��ȫѡ$day����ʱ��")

        # ȫѡ/ȡ��ȫѡ�¼�
        $chkDay.Add_CheckedChanged({
            param($sender)
            $targetDay = $sender.Text
            $checkedState = $sender.Checked
            # ����ʱ����б�ƥ�䵱ǰ�����յ�������
            for ($i = 0; $i -lt $lstLogonTime.Items.Count; $i++) {
                $item = $lstLogonTime.Items[$i].ToString()
                if ($item -like "$targetDay(*") {  # ƥ�䡰��һ(xx:xx-xx:xx)����ʽ
                    $lstLogonTime.SetItemChecked($i, $checkedState)
                }
            }
        })

        $weekDayTable.Controls.Add($chkDay)
        $script:weekDayCheckboxes += $chkDay
    }

    $mainTable.Controls.Add($weekDayTable, 0, 1)

    # 7. ��¼ʱ��ѡ���б������һ�µ�ʱ��θ�ʽ��
    $lstLogonTime = New-Object System.Windows.Forms.CheckedListBox
    $lstLogonTime.Dock = "Fill"
    $lstLogonTime.IntegralHeight = $false
    $lstLogonTime.ScrollAlwaysVisible = $true
    $lstLogonTime.CheckOnClick = $true
    $lstLogonTime.Font = New-Object System.Drawing.Font($lstLogonTime.Font.FontFamily, 8.5)

    # ����ʱ��Σ���ʽ����һ(0:00-1:00)������(23:00-0:00)��
    foreach ($day in $script:weekDaysUI) {
        for ($hour = 0; $hour -lt 24; $hour++) {
            # ����ʱ��ν���Сʱ�����촦��23�����Ϊ0�㣩
            $endHour = if ($hour -eq 23) { 0 } else { $hour + 1 }
            # ��ʽ��ʱ���
            $timeSlot = "$day($hour`:00-$endHour`:00)"
            $lstLogonTime.Items.Add($timeSlot, $false)
        }
    }

    $listPanel = New-Object System.Windows.Forms.Panel
    $listPanel.Dock = "Fill"
    $listPanel.Controls.Add($lstLogonTime)
    $mainTable.Controls.Add($listPanel, 0, 2)

    # 8. �ײ���ť���
    $bottomBtnPanel = New-Object System.Windows.Forms.Panel
    $bottomBtnPanel.Dock = "Fill"
    $bottomBtnPanel.Padding = 5

    $btnSaveTime = New-Object System.Windows.Forms.Button
    $btnSaveTime.Text = "������������"
    $btnSaveTime.Location = New-Object System.Drawing.Point(230, 10)
    $btnSaveTime.Width = 120
    $btnSaveTime.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $btnSaveTime.ForeColor = [System.Drawing.Color]::White
    $btnSaveTime.FlatStyle = "Flat"
    $toolTip.SetToolTip($btnSaveTime, "����ѡʱ��Ӧ�õ�ѡ���û�")

    $btnCancelTime = New-Object System.Windows.Forms.Button
    $btnCancelTime.Text = "ȡ��"
    $btnCancelTime.Location = New-Object System.Drawing.Point(370, 10)
    $btnCancelTime.Width = 120
    $btnCancelTime.BackColor = [System.Drawing.Color]::FromArgb(169, 169, 169)
    $btnCancelTime.ForeColor = [System.Drawing.Color]::White
    $btnCancelTime.FlatStyle = "Flat"

    $bottomBtnPanel.Controls.Add($btnSaveTime)
    $bottomBtnPanel.Controls.Add($btnCancelTime)
    $mainTable.Controls.Add($bottomBtnPanel, 0, 3)

    # ---------------------- 9. ��������1��AD logonHours �� UIʱ��� ----------------------
    function Convert-LogonHoursToTimeSlots {
        param([byte[]]$logonHours)
        $timeSlots = @()

        # ����Ĭ��ֵ��δ����ʱ��������ʱ�䣩
        if (-not $logonHours -or $logonHours.Length -ne 21) {
            $logonHours = [byte[]]::CreateInstance([byte], 21)
            for ($i = 0; $i -lt 21; $i++) { $logonHours[$i] = 0xFF }
        }

        # ���屾��ʱ�����ڡ�UI���ڵ�ӳ��
        $localWeekDayToUI = @{
            0 = "����"
            1 = "��һ"
            2 = "�ܶ�"
            3 = "����"
            4 = "����"
            5 = "����"
            6 = "����"
        }

        # ����AD��21�ֽڣ�7���3�ֽ�/�죩
        for ($byteIndex = 0; $byteIndex -lt 21; $byteIndex++) {
            $currentByte = $logonHours[$byteIndex]
            # 1. ���㵱ǰ�ֽڶ�Ӧ��AD����������0-6��
            $adWeekDayIndex = [math]::Floor($byteIndex / 3)
            # 2. ���㵱ǰ�ֽڶ�Ӧ��UTCСʱ�Σ�0-7,8-15,16-23��
            $hourSegment = $byteIndex % 3
            $hourBase = $hourSegment * 8

            # ����ÿ��Bit��Ӧ��UTCСʱ
            for ($bitIndex = 0; $bitIndex -lt 8; $bitIndex++) {
                $utcHour = $hourBase + $bitIndex
                # 3. ����UTCʱ�䣨ʹ�ù̶�����ȷ�����ڼ�����ȷ��
                $utcTime = [DateTime]::new(2023, 1, 1 + $adWeekDayIndex, $utcHour, 0, 0, [DateTimeKind]::Utc)
                # 4. ת��Ϊ����ʱ��
                $localTime = $utcTime.ToLocalTime()
                $localHour = $localTime.Hour
                $endHour = if ($localHour -eq 23) { 0 } else { $localHour + 1 }
                
                # 5. ��ȡ����ʱ�����������
                $localWeekDayNum = $localTime.DayOfWeek.value__
                # 6. ӳ�䵽UI��ʾ����������
                $uiWeekDay = $localWeekDayToUI[$localWeekDayNum]

                # 7. ����ǰBitΪ1�������¼�������ʱ���
                if (($currentByte -band (1 -shl $bitIndex)) -ne 0) {
                    $timeSlot = "$uiWeekDay($localHour`:00-$endHour`:00)"
                    $timeSlots += $timeSlot
                }
            }
        }
        return $timeSlots
    }

    # ---------------------- 10. ��������2��UIʱ��� �� AD logonHours�������޸����� ----------------------
    function Convert-TimeSlotsToLogonHours {
        param([string[]]$selectedTimeSlots)
        $logonHours = New-Object byte[] 21
        
        if (-not $selectedTimeSlots -or $selectedTimeSlots.Count -eq 0) {
            return $logonHours
        }

        # ����UI���ڡ�����ʱ������������ӳ��
        $uiToLocalWeekDay = @{
            "��һ" = 1
            "�ܶ�" = 2
            "����" = 3
            "����" = 4
            "����" = 5
            "����" = 6
            "����" = 0
        }

        foreach ($timeSlot in $selectedTimeSlots) {
            if ($timeSlot -match '^(��һ|�ܶ�|����|����|����|����|����)\((\d+):00-(\d+):00\)$') {
                $uiWeekDay = $matches[1]
                $localHour = [int]$matches[2]
                $endHour = [int]$matches[3]
                
                # 1. ��ȡ����ʱ�����������
                $localWeekDayNum = $uiToLocalWeekDay[$uiWeekDay]
                
                # 2. ��������ʱ�䣨ʹ�ù̶�����ȷ�����ڼ�����ȷ��
                $localTime = [DateTime]::new(2023, 1, 1 + $localWeekDayNum, $localHour, 0, 0, [DateTimeKind]::Local)
                
                # 3. ת��ΪUTCʱ��
                $utcTime = $localTime.ToUniversalTime()
                $utcHour = $utcTime.Hour
                
                # 4. ����AD����������0-6��
                $adWeekDayIndex = $utcTime.DayOfWeek.value__
                
                # 5. �����ֽ�λ�ú�λ����
                $hourSegment = [math]::Floor($utcHour / 8)
                $byteIndex = $adWeekDayIndex * 3 + $hourSegment
                $bitIndex = $utcHour % 8

                # 6. ��֤�ֽ������Ϸ���
                if ($byteIndex -ge 0 -and $byteIndex -lt 21) {
                    # 7. ��λ�����¼
                    $logonHours[$byteIndex] = $logonHours[$byteIndex] -bor (1 -shl $bitIndex)
                }
            }
        }
        return $logonHours
    }

    # 11. �����û����е�¼ʱ������
    function LoadUserLogonTime {
        try {
            $firstUserLogonHours = $null
            $hasDifferentSettings = $false
            $selectedSamAccounts = $selectedUsers | ForEach-Object { $_.Cells["SamAccountName"].Value }

            # ��ȡ��һ���û��ĵ�¼ʱ����Ϊ��׼
            foreach ($row in $selectedUsers) {
                $samAccountName = $row.Cells["SamAccountName"].Value
                $displayName = $row.Cells["DisplayName"].Value
                $displayName = if ([string]::IsNullOrEmpty($displayName)) { $samAccountName } else { $displayName }
                
                # Զ�̲�ѯ�û�logonHours
                $userLogonHours = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                    param($sam)
                    Import-Module ActiveDirectory -ErrorAction Stop
                    $user = Get-ADUser -Identity $sam -Properties logonHours -ErrorAction Stop
                    return $user.logonHours
                } -ArgumentList $samAccountName -ErrorAction Stop

                # �Աȶ��û������Ƿ�һ��
                if (-not $firstUserLogonHours) {
                    $firstUserLogonHours = $userLogonHours
                } else {
                    $isSame = $true
                    if (($firstUserLogonHours -eq $null) -xor ($userLogonHours -eq $null)) {
                        $isSame = $false
                    } elseif ($firstUserLogonHours -ne $null) {
                        if ($firstUserLogonHours.Length -ne $userLogonHours.Length) { $isSame = $false }
                        else {
                            for ($i = 0; $i -lt $firstUserLogonHours.Length; $i++) {
                                if ($firstUserLogonHours[$i] -ne $userLogonHours[$i]) { $isSame = $false; break }
                            }
                        }
                    }
                    if (-not $isSame) {
                        $hasDifferentSettings = $true
                    }
                }
            }

            # ת��Ϊ�¸�ʽ��ʱ����б�
            $allowedTimeSlots = Convert-LogonHoursToTimeSlots -logonHours $firstUserLogonHours
            $isAllowAll = $firstUserLogonHours -and $firstUserLogonHours.Where({ $_ -eq 0xFF }).Count -eq 21

            # ����UI״̬
            $chkAllowAllTime.Checked = $isAllowAll
            $lstLogonTime.Enabled = -not $isAllowAll
            $script:weekDayCheckboxes | ForEach-Object { $_.Enabled = -not $isAllowAll }

            # ��ѡ���������ʱ���
            $lstLogonTime.ClearSelected()
            for ($i = 0; $i -lt $lstLogonTime.Items.Count; $i++) {
                $item = $lstLogonTime.Items[$i].ToString()
                if ($allowedTimeSlots -contains $item) {
                    $lstLogonTime.SetItemChecked($i, $true)
                }
            }

            # ͬ�����¹�����ȫѡ��״̬
            foreach ($day in $script:weekDaysUI) {
                $dayItems = $lstLogonTime.Items | Where-Object { $_ -like "$day(*" }
                $checkedDayItems = $dayItems | Where-Object { $lstLogonTime.GetItemChecked($lstLogonTime.Items.IndexOf($_)) }
                $chkDay = $script:weekDayCheckboxes | Where-Object { $_.Text -eq $day }
                $chkDay.Checked = ($checkedDayItems.Count -eq $dayItems.Count) -and ($dayItems.Count -gt 0)
            }

            # ��ʾ���û����ò�һ��
            if ($hasDifferentSettings) {
                [System.Windows.Forms.MessageBox]::Show(
                    "ѡ�е��û����е�¼ʱ�����ò�һ�£���ͳһ������ѡ����������", 
                    "��ʾ", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("���ص�¼ʱ��ʧ�ܣ�$($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $restrictTimeForm.Close()
        }
    }

    # 12. ���ڼ���ʱִ��
    $restrictTimeForm.Add_Shown({ 
        LoadUserLogonTime 
    })

    # 13. �����¼ʱ������
    $btnSaveTime.Add_Click({
        $isAllowAll = $chkAllowAllTime.Checked
        $selectedTimeSlots = @()
        if (-not $isAllowAll) {
            # �ռ�UIѡ�е�ʱ��Σ��¸�ʽ��
            for ($i = 0; $i -lt $lstLogonTime.Items.Count; $i++) {
                if ($lstLogonTime.GetItemChecked($i)) {
                    $selectedTimeSlots += $lstLogonTime.Items[$i].ToString()
                }
            }
        }

        # ����ȷ����Ϣ
        if ($isAllowAll) {
            $confirmMsg = "��Ϊ $($selectedUsers.Count) ���û����á���������ʱ���¼��`n`nȷ��������"
        } elseif ($selectedTimeSlots.Count -eq 0) {
            $confirmMsg = "��Ϊ $($selectedUsers.Count) ���û����á���ֹ����ʱ���¼��`n`nȷ��������"
        } else {
            $confirmMsg = "��Ϊ $($selectedUsers.Count) ���û����������¼��ʱ�䣺`n"
            foreach ($day in $script:weekDaysUI) {
                $daySlots = $selectedTimeSlots | Where-Object { $_ -like "$day(*" }
                if ($daySlots) {
                    $confirmMsg += "$day��$($daySlots -join "��")`n"
                }
            }
            $confirmMsg += "`nȷ��������"
        }

        # ȷ�ϱ���
        if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, "ȷ�ϱ���", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -ne "Yes") {
            return
        }

        try {
            $successUsers = @()
            $failUsers = @()

            # ����AD��Ҫ��logonHours�ֽ�����
            $logonHours = if ($isAllowAll) {
                # ��������ʱ�䣺21�ֽ�ȫΪ0xFF
                $bytes = New-Object byte[] 21
                for ($i = 0; $i -lt 21; $i++) { $bytes[$i] = 0xFF }
                $bytes
            } else {
                Convert-TimeSlotsToLogonHours -selectedTimeSlots $selectedTimeSlots
            }

            # ����Ӧ�õ�ѡ���û�
            foreach ($row in $selectedUsers) {
                $displayName = $row.Cells["DisplayName"].Value
                $samAccountName = $row.Cells["SamAccountName"].Value
                $displayName = if ([string]::IsNullOrEmpty($displayName)) { "δ������ʾ��" } else { $displayName }
                $userLabel = "$displayName��$samAccountName��"

                try {
                    $bytesToSet = [byte[]]$logonHours

                    # Զ��ִ�б��棬���Զ��ִ�н��
                    $remoteResult = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                        param($sam, $hours)
                        Import-Module ActiveDirectory -ErrorAction Stop
                        $hoursBytes = [byte[]]$hours
                        
                        Set-ADUser -Identity $sam -Replace @{logonHours = $hoursBytes} -ErrorAction Stop
                        
                        # ��֤�����ú�������ѯ��ȷ���Ƿ���Ч
                        $updatedUser = Get-ADUser -Identity $sam -Properties logonHours -ErrorAction Stop
                        $isUpdated = $true
                        if ($updatedUser.logonHours.Length -ne $hoursBytes.Length) {
                            $isUpdated = $false
                        } else {
                            for ($i = 0; $i -lt $hoursBytes.Length; $i++) {
                                if ($updatedUser.logonHours[$i] -ne $hoursBytes[$i]) {
                                    $isUpdated = $false; break
                                }
                            }
                        }
                        return $isUpdated
                    } -ArgumentList $samAccountName, $bytesToSet -ErrorAction Stop

                    if ($remoteResult) {
                        $successUsers += $userLabel
                    } else {
                        $failUsers += "$userLabel - Զ�����ú���֤��һ��"
                    }
                } catch {
                    $failMsg = $_.Exception.Message
                    $failUsers += "$userLabel - $failMsg"
                }
            }

            # ��ʾ���
            $resultTitle = if ($failUsers.Count -eq 0) { "�ɹ�" } else { "���" }
            $resultIcon = if ($failUsers.Count -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }

            $resultMsg = "��¼ʱ������������ɣ�`n"
            $resultMsg += "`n�ɹ��û����� $($successUsers.Count) ������"
            $resultMsg += if ($successUsers.Count -eq 0) { "`n��" } else { "`n$($successUsers -join "`n")" }
            $resultMsg += "`n`nʧ���û����� $($failUsers.Count) ������"
            $resultMsg += if ($failUsers.Count -eq 0) { "`n��" } else { "`n$($failUsers -join "`n")" }

            [System.Windows.Forms.MessageBox]::Show($resultMsg, $resultTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

            # ȫ���ɹ���رմ���
            if ($failUsers.Count -eq 0) {
                $restrictTimeForm.Close()
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("�����쳣��$($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # 14. ȡ����ť�߼�
    $btnCancelTime.Add_Click({
        $restrictTimeForm.Close()
    })

    # 15. ���ش���
    $restrictTimeForm.Controls.Add($mainTable)
    $restrictTimeForm.ShowDialog() | Out-Null
}
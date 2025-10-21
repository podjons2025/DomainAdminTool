<# 
���ĺ������������Ƶ�¼���������
#>

function ShowRestrictLoginForm {
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("�������ӵ����", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # 1. ����Ƿ�ѡ���û���֧�ֶ��û���
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
	
	
    # ��������Ա���У��
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

    # 2. �����´���
    $restrictForm = New-Object System.Windows.Forms.Form
    $restrictForm.Text = "�����û���¼�����(֧���˺Ŷ�ѡ)"
    $restrictForm.Size = New-Object System.Drawing.Size(800, 550)  # �߶�����50��Ӧ������
    $restrictForm.StartPosition = "CenterParent"
    $restrictForm.FormBorderStyle = "FixedDialog"
    $restrictForm.MaximizeBox = $false
    $restrictForm.MinimizeBox = $false

    # ����ToolTip���������ʾ��ť��ʾ
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 5000  # ��ʾ��ʾʱ�䣨���룩
    $toolTip.InitialDelay = 1000  # �����ͣ���ӳ���ʾʱ��
    $toolTip.ReshowDelay = 500    # ������ʾ��ʾ���ӳ�ʱ��
    $toolTip.ShowAlways = $true   # ��ʹ���ڲ��ڽ���Ҳ��ʾ

    # �洢��ʾ�ı���ԭʼ��������ӳ�䣨�ؼ���
    $script:computerNameMap = @{}
    # �洢ԭʼ�б����ݣ������������ˣ�
    $script:allowedOriginalItems = @()
    $script:allComputersOriginalItems = @()

    # 3. ���ڲ��֣�ʹ��TableLayoutPanel�Ű�
    $mainTable = New-Object System.Windows.Forms.TableLayoutPanel
    $mainTable.Dock = "Fill"
    $mainTable.Padding = 10
    $mainTable.ColumnCount = 5  # ���б�(2��) + ��ť��(1��) + ���б�(2��)
    $mainTable.RowCount = 2     # �б���(1��) + ��ť��(1��)
    $mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
    $mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 10)))  # ���
    $mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 65))) # ��ť��
    $mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 10)))  # ���
    $mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 85)))
    $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))

    #4. ����б������¼�ļ�����������������ܣ�
    # �ع�ΪTableLayoutPanel���ڲ��ֱ��⡢�������б�����ͳ��
    $leftPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $leftPanel.Dock = "Fill"
    $leftPanel.RowCount = 4  # �����С������С��б��С�����ͳ����
    $leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 25)))  # ����߶�
    $leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))  # ������߶�
    $leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # �б�ռ��
    $leftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))  # ͳ�Ƹ߶�

    # �����ǩ
    $lblAllowed = New-Object System.Windows.Forms.Label
    $lblAllowed.Text = "�����¼�ļ������ѡ���û����ã�"
    $lblAllowed.Dock = "Fill"
    $lblAllowed.Font = New-Object System.Drawing.Font($lblAllowed.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $leftPanel.Controls.Add($lblAllowed, 0, 0)

    # ����������
    $searchAllowedPanel = New-Object System.Windows.Forms.Panel
    $searchAllowedPanel.Dock = "Fill"
    $searchAllowedPanel.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)

    $lblSearchAllowed = New-Object System.Windows.Forms.Label
    $lblSearchAllowed.Text = "������"
    $lblSearchAllowed.Location = New-Object System.Drawing.Point(0, 5)
    $lblSearchAllowed.AutoSize = $true

    $txtSearchAllowed = New-Object System.Windows.Forms.TextBox
    $txtSearchAllowed.Dock = "Fill"
    $txtSearchAllowed.Margin = New-Object System.Windows.Forms.Padding(35, 0, 0, 0)  # ������ǩλ��
    $toolTip.SetToolTip($txtSearchAllowed, "֧��ģ�������������ִ�Сд")
    # ģ��ռλ�ı�
    $allowedPlaceholder = "������������IP����..."
    $txtSearchAllowed.Text = $allowedPlaceholder
    $txtSearchAllowed.ForeColor = [System.Drawing.Color]::Gray  # ռλ�ı���ɫ

    # �󶨽����¼�ʵ��ռλ�ı�Ч��
    $txtSearchAllowed.Add_GotFocus({
        if ($this.Text -eq $allowedPlaceholder) {
            $this.Text = ""
            $this.ForeColor = [System.Drawing.Color]::Black  # �����ı���ɫ
        }
    })
    $txtSearchAllowed.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($this.Text)) {
            $this.Text = $allowedPlaceholder
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })

    $searchAllowedPanel.Controls.Add($txtSearchAllowed)
    $searchAllowedPanel.Controls.Add($lblSearchAllowed)
    $leftPanel.Controls.Add($searchAllowedPanel, 0, 1)

    # �б��
    $lstAllowed = New-Object System.Windows.Forms.ListBox
    $lstAllowed.Dock = "Fill"
    $lstAllowed.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended  # ֧�ֶ�ѡ
    $lstAllowed.IntegralHeight = $false  # �����Զ��߶ȣ���Ӧ��壩
    $lstAllowed.ScrollAlwaysVisible = $true  # ʼ����ʾ������
    $leftPanel.Controls.Add($lstAllowed, 0, 2)

    # ����ͳ�Ʊ�ǩ���ײ���
    $lblAllowedCount = New-Object System.Windows.Forms.Label
    $lblAllowedCount.Text = "�����¼������� 0 ̨"
    $lblAllowedCount.Dock = "Fill"
    $lblAllowedCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblAllowedCount.Font = New-Object System.Drawing.Font($lblAllowedCount.Font.FontFamily, 9)
    $leftPanel.Controls.Add($lblAllowedCount, 0, 3)

    $mainTable.Controls.Add($leftPanel, 0, 0)  # ���б���ڵ�0�е�0��

    # 5. �м䰴ť���б��ƶ����ƣ�����λ������ֱ���У�
    # ʹ��FlowLayoutPanel���ư�ť����
    $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnPanel.Dock = "Fill"
    # ���󶥲��ڱ߾�ʹ��ť�������м䣨ԭ30����Ϊ120�����ݴ��ڸ߶ȼ��㣩
    $btnPanel.Padding = New-Object System.Windows.Forms.Padding(0, 50, 0, 0)  # �����������ӣ�ʹ��ť����
    $btnPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown  # ��ť��ֱ����
    $btnPanel.WrapContents = $false  # ���Զ�����
    $btnPanel.AutoSize = $false  # ���Զ�������С

    # ��ť���֣���ֱ���У����ּ��
    $btnAddAll = New-Object System.Windows.Forms.Button
    $btnAddAll.Text = "<<"
    $btnAddAll.Width = 60  # �̶���ť���
    $btnAddAll.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)  # ���¸�8���ؼ��
    $toolTip.SetToolTip($btnAddAll, "����������ڼ�����������б�")

    $btnAddSelected = New-Object System.Windows.Forms.Button
    $btnAddSelected.Text = "<"
    $btnAddSelected.Width = 60
    $btnAddSelected.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)
    $toolTip.SetToolTip($btnAddSelected, "���ѡ�еļ�����������б�")

    $btnRemoveSelected = New-Object System.Windows.Forms.Button
    $btnRemoveSelected.Text = ">"
    $btnRemoveSelected.Width = 60
    $btnRemoveSelected.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)
    $toolTip.SetToolTip($btnRemoveSelected, "�������б��Ƴ�ѡ�еļ����")

    $btnRemoveAll = New-Object System.Windows.Forms.Button
    $btnRemoveAll.Text = ">>"
    $btnRemoveAll.Width = 60
    $btnRemoveAll.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)
    $toolTip.SetToolTip($btnRemoveAll, "��������б�")

    # ��˳����Ӱ�ť
    $btnPanel.Controls.Add($btnAddAll)
    $btnPanel.Controls.Add($btnAddSelected)
    $btnPanel.Controls.Add($btnRemoveSelected)
    $btnPanel.Controls.Add($btnRemoveAll)
    $mainTable.Controls.Add($btnPanel, 2, 0)  # ��ť���ڵ�2�е�0��

    # 6. �Ҳ��б��������м����
    # �ع�ΪTableLayoutPanel���ڲ��ֱ��⡢�������б�����ͳ��
    $rightPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $rightPanel.Dock = "Fill"
    $rightPanel.RowCount = 4  # �����С������С��б��С�����ͳ����
    $rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 25)))  # ����߶�
    $rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))  # ������߶�
    $rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # �б�ռ��
    $rightPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))  # ͳ�Ƹ߶�

    # �����ǩ
    $lblAllComputers = New-Object System.Windows.Forms.Label
    $lblAllComputers.Text = "�������м����������ؼ�����"
    $lblAllComputers.Dock = "Fill"
    $lblAllComputers.Font = New-Object System.Drawing.Font($lblAllComputers.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $rightPanel.Controls.Add($lblAllComputers, 0, 0)

    # ����������
    $searchAllPanel = New-Object System.Windows.Forms.Panel
    $searchAllPanel.Dock = "Fill"
    $searchAllPanel.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)

    $lblSearchAll = New-Object System.Windows.Forms.Label
    $lblSearchAll.Text = "������"
    $lblSearchAll.Location = New-Object System.Drawing.Point(0, 5)
    $lblSearchAll.AutoSize = $true

    $txtSearchAll = New-Object System.Windows.Forms.TextBox
    $txtSearchAll.Dock = "Fill"
    $txtSearchAll.Margin = New-Object System.Windows.Forms.Padding(35, 0, 0, 0)  # ������ǩλ��
    $toolTip.SetToolTip($txtSearchAll, "֧��ģ�������������ִ�Сд")
    # ģ��ռλ�ı�
    $allPlaceholder = "������������IP����..."
    $txtSearchAll.Text = $allPlaceholder
    $txtSearchAll.ForeColor = [System.Drawing.Color]::Gray  # ռλ�ı���ɫ

    # �󶨽����¼�ʵ��ռλ�ı�Ч��
    $txtSearchAll.Add_GotFocus({
        if ($this.Text -eq $allPlaceholder) {
            $this.Text = ""
            $this.ForeColor = [System.Drawing.Color]::Black  # �����ı���ɫ
        }
    })
    $txtSearchAll.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($this.Text)) {
            $this.Text = $allPlaceholder
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })

    $searchAllPanel.Controls.Add($txtSearchAll)
    $searchAllPanel.Controls.Add($lblSearchAll)
    $rightPanel.Controls.Add($searchAllPanel, 0, 1)

    # �б��
    $lstAllComputers = New-Object System.Windows.Forms.ListBox
    $lstAllComputers.Dock = "Fill"
    $lstAllComputers.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
    $lstAllComputers.IntegralHeight = $false
    $lstAllComputers.ScrollAlwaysVisible = $true
    $rightPanel.Controls.Add($lstAllComputers, 0, 2)

    # ����ͳ�Ʊ�ǩ���ײ���
    $lblAllComputersCount = New-Object System.Windows.Forms.Label
    $lblAllComputersCount.Text = "�б������� 0 ̨"
    $lblAllComputersCount.Dock = "Fill"
    $lblAllComputersCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblAllComputersCount.Font = New-Object System.Drawing.Font($lblAllComputersCount.Font.FontFamily, 9)
    $rightPanel.Controls.Add($lblAllComputersCount, 0, 3)

    $mainTable.Controls.Add($rightPanel, 4, 0)  # ���б���ڵ�4�е�0��

    # 7. �ײ���ť������/ȡ��
    $bottomBtnPanel = New-Object System.Windows.Forms.Panel
    $bottomBtnPanel.Dock = "Fill"
    $bottomBtnPanel.Padding = 5

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "������������"
    $btnSave.Location = New-Object System.Drawing.Point(250, 10)
    $btnSave.Width = 120
    $btnSave.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $btnSave.ForeColor = [System.Drawing.Color]::White
    $btnSave.FlatStyle = "Flat"

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "ȡ��"
    $btnCancel.Location = New-Object System.Drawing.Point(400, 10)
    $btnCancel.Width = 120
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(169, 169, 169)
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = "Flat"

    $bottomBtnPanel.Controls.Add($btnSave)
    $bottomBtnPanel.Controls.Add($btnCancel)
    $mainTable.Controls.Add($bottomBtnPanel, 0, 1)  # �ײ���ť���ڵ�0�е�1��
    $mainTable.SetColumnSpan($bottomBtnPanel, 5)  # ��5�У�ռ���ײ���ȣ�

    # 8. ����ͳ�Ƹ��º���
    function UpdateCounts {
        $lblAllowedCount.Text = "�����¼������� $($lstAllowed.Items.Count) ̨"
        $lblAllComputersCount.Text = "�б������� $($lstAllComputers.Items.Count) ̨"
    }

    # 9. �������˺���������ռλ�ı������
    function FilterList($sourceList, $targetList, $originalItems, $searchText, $placeholder) {
        # �����ռλ�ı�����Ϊ������
        $actualSearchText = if ($searchText -eq $placeholder) { "" } else { $searchText }
        
        $targetList.Items.Clear()
        if ([string]::IsNullOrWhiteSpace($actualSearchText)) {
            # ����Ϊ��ʱ��ʾ����ԭʼ��
            $originalItems | ForEach-Object { $targetList.Items.Add($_) | Out-Null }
        }
        else {
            # ģ��ƥ�䣨�����ִ�Сд��
            $lowerSearch = $actualSearchText.ToLower()
            $originalItems | Where-Object { $_.ToLower().Contains($lowerSearch) } | ForEach-Object {
                $targetList.Items.Add($_) | Out-Null
            }
        }
    }

    # 10. �������ڼ����
    function LoadDomainComputers {
        try {
            # ���ӳ����ԭʼ�б�
            $script:computerNameMap.Clear()
            $script:allowedOriginalItems = @()
            $script:allComputersOriginalItems = @()

            # Զ�̴���ػ�ȡ���м����������IP��ַ��
            $allComputers = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                Import-Module ActiveDirectory -ErrorAction Stop
                # ��ȡ�������õļ�������������ƺ�IP��ַ
                Get-ADComputer -Filter { Enabled -eq $true } -Properties Name, IPv4Address | 
                    Select-Object Name, IPv4Address | 
                    Sort-Object Name  # ����������
            } -ErrorAction Stop

            # �����������ݣ�����"������--��IP��"��ʽ����ʾ�ı�
            $displayTexts = @()
            foreach ($comp in $allComputers) {
                $hostName = $comp.Name
                $ipAddress = $comp.IPv4Address
                
                # ������IP��ַ�����
                $displayIp = if ($ipAddress) { $ipAddress } else { "��IP" }
                $displayText = "$hostName -- ($displayIp)"
                
                # ��ӵ���ʾ�б��ӳ���
                $displayTexts += $displayText
                $script:computerNameMap[$displayText] = $hostName  # ӳ����ʾ�ı���ԭʼ������
            }

            # �����Ҳ��б�ԭʼ���ݲ���ʼ����ʾ
            $script:allComputersOriginalItems = $displayTexts
            $lstAllComputers.Items.Clear()
            $displayTexts | ForEach-Object { $lstAllComputers.Items.Add($_) | Out-Null }

            # ����ѡ���û��ġ����������¼�������
            $firstUserWorkstations = $null
            $hasDifferentSettings = $false
            foreach ($row in $selectedUsers) {
                $samAccountName = $row.Cells["SamAccountName"].Value
                $userWorkstations = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                    param($sam)
                    Import-Module ActiveDirectory -ErrorAction Stop
                    Get-ADUser -Identity $sam -Properties userWorkstations -ErrorAction Stop | 
                        Select-Object -ExpandProperty userWorkstations
                } -ArgumentList $samAccountName -ErrorAction Stop

                # �״μ���ʱ��¼��һ���û�������
                if (-not $firstUserWorkstations) {
                    $firstUserWorkstations = $userWorkstations
                }
                # �����û������Ƿ�һ��
                elseif ($userWorkstations -ne $firstUserWorkstations) {
                    $hasDifferentSettings = $true
                }
            }

            # ��������б�����
            $allowedDisplayTexts = @()
            if ($hasDifferentSettings) {
                [System.Windows.Forms.MessageBox]::Show("ѡ�е��û����е�¼�������ò�һ�£���ͳһ������ѡ����������", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            elseif ($firstUserWorkstations) {
                # ������м�������������ŷָȥ�գ�
                $allowedHostNames = $firstUserWorkstations -split "," | 
                    ForEach-Object { $_.Trim() } | 
                    Where-Object { $_ -ne "" }

                # ת��Ϊ������--��IP����ʽ��ʾ
                foreach ($hostName in $allowedHostNames) {
                    # ���Ҷ�Ӧ����ʾ�ı�
                    $mappedText = $script:computerNameMap.GetEnumerator() | 
                        Where-Object { $_.Value -eq $hostName } | 
                        Select-Object -ExpandProperty Key -First 1

                    if ($mappedText) {
                        $allowedDisplayTexts += $mappedText
                    }
                    else {
                        # �����Ѳ����ڵļ�������������Ҳ�����
                        $customText = "$hostName--���Ѳ����ڣ�"
                        $allowedDisplayTexts += $customText
                        # ��ӵ�ӳ���ȷ������ʱ����ȷ��ȡԭʼ������
                        $script:computerNameMap[$customText] = $hostName
                    }
                }
            }

            # ��������б�ԭʼ���ݲ���ʼ����ʾ
            $script:allowedOriginalItems = $allowedDisplayTexts
            $lstAllowed.Items.Clear()
            $allowedDisplayTexts | ForEach-Object { $lstAllowed.Items.Add($_) | Out-Null }

            # ������ɺ��������ͳ��
            UpdateCounts

        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("���������ʧ�ܣ�$($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $restrictForm.Close()
        }
    }

    # 11. ���������¼�������ռλ�ı�������
    $txtSearchAllowed.Add_TextChanged({
        FilterList -sourceList $lstAllowed `
                   -targetList $lstAllowed `
                   -originalItems $script:allowedOriginalItems `
                   -searchText $txtSearchAllowed.Text `
                   -placeholder $allowedPlaceholder
        UpdateCounts
    })

    $txtSearchAll.Add_TextChanged({
        FilterList -sourceList $lstAllComputers `
                   -targetList $lstAllComputers `
                   -originalItems $script:allComputersOriginalItems `
                   -searchText $txtSearchAll.Text `
                   -placeholder $allPlaceholder
        UpdateCounts
    })

    # ���ڼ���ʱִ�м��������
    $restrictForm.Add_Shown({ LoadDomainComputers })

# 12. �б��ƶ���ť�߼�
    # ȫ����ӵ������б�
    $btnAddAll.Add_Click({
        # ����ԭʼ���ݲ�����������������Ӱ��
        foreach ($item in $script:allComputersOriginalItems) {
            if (-not $script:allowedOriginalItems.Contains($item)) {
                $script:allowedOriginalItems += $item
            }
        }
        # ˢ������б����ֵ�ǰ����״̬��
        FilterList -sourceList $lstAllowed `
                   -targetList $lstAllowed `
                   -originalItems $script:allowedOriginalItems `
                   -searchText $txtSearchAllowed.Text `
                   -placeholder $allowedPlaceholder
        UpdateCounts  # ��������ͳ��
    })

    # ѡ����ӵ������б�
    $btnAddSelected.Add_Click({
        $selectedItems = @($lstAllComputers.SelectedItems)
        foreach ($item in $selectedItems) {
            if (-not $script:allowedOriginalItems.Contains($item)) {
                $script:allowedOriginalItems += $item
            }
        }
        # ˢ������б����ֵ�ǰ����״̬��
        FilterList -sourceList $lstAllowed `
                   -targetList $lstAllowed `
                   -originalItems $script:allowedOriginalItems `
                   -searchText $txtSearchAllowed.Text `
                   -placeholder $allowedPlaceholder
        $lstAllComputers.ClearSelected()  # ȡ���Ҳ�ѡ��
        UpdateCounts  # ��������ͳ��
    })

    # ѡ�д������б��Ƴ�
    $btnRemoveSelected.Add_Click({
        $selectedItems = @($lstAllowed.SelectedItems)
        foreach ($item in $selectedItems) {
            $script:allowedOriginalItems = $script:allowedOriginalItems | Where-Object { $_ -ne $item }
        }
        # ˢ������б����ֵ�ǰ����״̬��
        FilterList -sourceList $lstAllowed `
                   -targetList $lstAllowed `
                   -originalItems $script:allowedOriginalItems `
                   -searchText $txtSearchAllowed.Text `
                   -placeholder $allowedPlaceholder
        $lstAllowed.ClearSelected()  # ȡ�����ѡ��
        UpdateCounts  # ��������ͳ��
    })

    # ��������б�
    $btnRemoveAll.Add_Click({
        if ([System.Windows.Forms.MessageBox]::Show("ȷ��Ҫ������������¼�ļ������", "ȷ��", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq "Yes") {
            $script:allowedOriginalItems = @()
            # ˢ������б����ֵ�ǰ����״̬��
            FilterList -sourceList $lstAllowed `
                       -targetList $lstAllowed `
                       -originalItems $script:allowedOriginalItems `
                       -searchText $txtSearchAllowed.Text `
                       -placeholder $allowedPlaceholder
            UpdateCounts  # ��������ͳ��
        }
    })
    

# 13. �����������ã���ȡԭʼ��������
    $btnSave.Add_Click({
        # 1. ��ԭʼ�����б���ȡ������������������������Ӱ�죩
        $allowedHostNames = @()
        foreach ($displayItem in $script:allowedOriginalItems) {
            # ��ӳ����ȡԭʼ������
            if ($script:computerNameMap.ContainsKey($displayItem)) {
                $allowedHostNames += $script:computerNameMap[$displayItem]
            }
            else {
                # ����ʾ�ı�����ȡ������������ʽ��֣�
                $hostNamePart = $displayItem -split "--", 2 | Select-Object -First 1
                $allowedHostNames += $hostNamePart.Trim()
            }
        }
        $allowedComputers = $allowedHostNames -join ","
        $isClearRestriction = [string]::IsNullOrEmpty($allowedComputers)  # �ж��Ƿ�Ϊ���������ơ�

        # 2. ȷ�ϱ���
        $confirmMsg = if ($isClearRestriction) {
            "��Ϊ $($selectedUsers.Count) ���û��������¼���ơ��������¼���м������`n`nȷ��������"
        } else {
            "��Ϊ $($selectedUsers.Count) ���û����������¼�ļ������`n$allowedComputers`n`nȷ��������"
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            $confirmMsg, 
            "ȷ�ϱ���", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne "Yes") { return }

        try {
            # ��¼�ɹ�/ʧ�ܵ��û�����
            $successUsers = @()  # ��ʽ��"��ʾ�����˺�����"
            $failUsers = @()     # ��ʽ��"��ʾ�����˺�����- ʧ��ԭ��"

            # ����ÿ��ѡ���û�����������
            foreach ($row in $selectedUsers) {
                $displayName = $row.Cells["DisplayName"].Value
                $samAccountName = $row.Cells["SamAccountName"].Value
                $displayName = if ([string]::IsNullOrEmpty($displayName)) { "δ������ʾ��" } else { $displayName }
                $userLabel = "$displayName��$samAccountName��"

                try {
                    Invoke-Command -Session $script:remoteSession -ScriptBlock {
                        param($sam, $workstations, $isClear)
                        Import-Module ActiveDirectory -ErrorAction Stop

                        if ($isClear) {
                            # �������ƣ����userWorkstations����
                            Set-ADUser -Identity $sam -Clear userWorkstations -ErrorAction Stop
                        } else {
                            # �������ƣ�����userWorkstations����
                            Set-ADUser -Identity $sam -Replace @{userWorkstations = $workstations} -ErrorAction Stop
                        }
                    } -ArgumentList $samAccountName, $allowedComputers, $isClearRestriction -ErrorAction Stop

                    $successUsers += $userLabel
                }
                catch {
                    $failReason = $_.Exception.Message
                    $failUsers += "$userLabel - $failReason"
                }
            }

            # 3. ���ɽ����ʾ
            $resultTitle = if ($failUsers.Count -eq 0) { "�ɹ�" } else { "���" }
            $resultIcon = if ($failUsers.Count -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
            
            $resultMsg = "������ɣ�`n"
            $resultMsg += "`n�ɹ��û����� $($successUsers.Count) ������"
            if ($successUsers.Count -eq 0) {
                $resultMsg += "`n��"
            } else {
                $resultMsg += "`n$($successUsers -join "`n")"
            }
            $resultMsg += "`n`nʧ���û����� $($failUsers.Count) ������"
            if ($failUsers.Count -eq 0) {
                $resultMsg += "`n��"
            } else {
                $resultMsg += "`n$($failUsers -join "`n")"
            }

            [System.Windows.Forms.MessageBox]::Show($resultMsg, $resultTitle, [System.Windows.Forms.MessageBoxButtons]::OK, $resultIcon)

            # ȫ���ɹ�ʱ�Զ��رմ���
            if ($failUsers.Count -eq 0) {
                $restrictForm.Close()
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("���������쳣��$($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    


    # 14. ȡ����ť�߼�
    $btnCancel.Add_Click({
        $restrictForm.Close()
    })

    # 15. ������ʾ
    $restrictForm.Controls.Add($mainTable)
    $restrictForm.ShowDialog() | Out-Null

    # ����ű�������
    Remove-Variable -Name computerNameMap -Scope Script -ErrorAction SilentlyContinue
    Remove-Variable -Name allowedOriginalItems -Scope Script -ErrorAction SilentlyContinue
    Remove-Variable -Name allComputersOriginalItems -Scope Script -ErrorAction SilentlyContinue
}
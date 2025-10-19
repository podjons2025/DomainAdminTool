<# 
���ĺ����������������Ƶ�¼��������ڣ�˫�б�棩
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
	
	
    # ---------------------- ��������������Ա���У�� ----------------------
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
    # ---------------------- �������� ----------------------	

    # 2. �����´���
    $restrictForm = New-Object System.Windows.Forms.Form
    $restrictForm.Text = "�����û���¼�����(֧���˺Ŷ�ѡ)"
    $restrictForm.Size = New-Object System.Drawing.Size(800, 500)
    $restrictForm.StartPosition = "CenterParent"
    $restrictForm.FormBorderStyle = "FixedDialog"
    $restrictForm.MaximizeBox = $false
    $restrictForm.MinimizeBox = $false

    # ����������ToolTip���������ʾ��ť��ʾ
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 5000  # ��ʾ��ʾʱ�䣨���룩
    $toolTip.InitialDelay = 1000  # �����ͣ���ӳ���ʾʱ��
    $toolTip.ReshowDelay = 500    # ������ʾ��ʾ���ӳ�ʱ��
    $toolTip.ShowAlways = $true   # ��ʹ���ڲ��ڽ���Ҳ��ʾ

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

    # ---------------------- 4. ����б������¼�ļ���� ----------------------
    $leftPanel = New-Object System.Windows.Forms.Panel
    $leftPanel.Dock = "Fill"

    $lblAllowed = New-Object System.Windows.Forms.Label
    $lblAllowed.Text = "�����¼�ļ������ѡ���û����ã�"
    $lblAllowed.Dock = "Top"
    $lblAllowed.Font = New-Object System.Drawing.Font($lblAllowed.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)

    $lstAllowed = New-Object System.Windows.Forms.ListBox
    $lstAllowed.Dock = "Fill"
    $lstAllowed.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended  # ֧�ֶ�ѡ
    $lstAllowed.IntegralHeight = $false  # �����Զ��߶ȣ���Ӧ��壩
    $lstAllowed.ScrollAlwaysVisible = $true  # ʼ����ʾ������

    $leftPanel.Controls.Add($lstAllowed)
    $leftPanel.Controls.Add($lblAllowed)  # ��ǩ���б��Ϸ�
    $mainTable.Controls.Add($leftPanel, 0, 0)  # ���б���ڵ�0�е�0��

    # ---------------------- 5. �м䰴ť���б��ƶ�����----------------------
    $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnPanel.Dock = "Fill"
    $btnPanel.Padding = New-Object System.Windows.Forms.Padding(0, 30, 0, 0)  # ������30���ؿհף�ʹ��ť��������
    $btnPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown  # ��ť��ֱ����
    $btnPanel.WrapContents = $false  # ���Զ�����
    $btnPanel.AutoSize = $false  # ���Զ�������С

    # ��ť���֣���ֱ���У����Ӽ��
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

    # ��˳����Ӱ�ť��FlowDirection=TopDown���ȼӵ������棩
    $btnPanel.Controls.Add($btnAddAll)
    $btnPanel.Controls.Add($btnAddSelected)
    $btnPanel.Controls.Add($btnRemoveSelected)
    $btnPanel.Controls.Add($btnRemoveAll)
    $mainTable.Controls.Add($btnPanel, 2, 0)  # ��ť���ڵ�2�е�0��

    # ---------------------- 6. �Ҳ��б��������м���� ----------------------
    $rightPanel = New-Object System.Windows.Forms.Panel
    $rightPanel.Dock = "Fill"

    $lblAllComputers = New-Object System.Windows.Forms.Label
    $lblAllComputers.Text = "�������м����������ؼ�����"
    $lblAllComputers.Dock = "Top"
    $lblAllComputers.Font = New-Object System.Drawing.Font($lblAllComputers.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)

    $lstAllComputers = New-Object System.Windows.Forms.ListBox
    $lstAllComputers.Dock = "Fill"
    $lstAllComputers.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
    $lstAllComputers.IntegralHeight = $false
    $lstAllComputers.ScrollAlwaysVisible = $true

    $rightPanel.Controls.Add($lstAllComputers)
    $rightPanel.Controls.Add($lblAllComputers)
    $mainTable.Controls.Add($rightPanel, 4, 0)  # ���б���ڵ�4�е�0��

    # ---------------------- 7. �ײ���ť������/ȡ�� ----------------------
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

    # ---------------------- 8. �������ڼ�����������߼��� ----------------------
    function LoadDomainComputers {
        try {
            # Զ�̴���ػ�ȡ���м������ֻȡ���������ȥ�أ�
            $allComputers = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                Import-Module ActiveDirectory -ErrorAction Stop
                # ��ȡ�������õļ�������ų��ظ�����Nameȥ�أ�
                Get-ADComputer -Filter { Enabled -eq $true } -Properties Name | 
                    Select-Object -ExpandProperty Name -Unique | 
                    Sort-Object  # ����������
            } -ErrorAction Stop

            # ����Ҳ��б�
            $lstAllComputers.Items.Clear()
            $allComputers | ForEach-Object { $lstAllComputers.Items.Add($_) | Out-Null }

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

            # �������б�
            $lstAllowed.Items.Clear()
            if ($hasDifferentSettings) {
                [System.Windows.Forms.MessageBox]::Show("ѡ�е��û����е�¼�������ò�һ�£���ͳһ������ѡ����������", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            elseif ($firstUserWorkstations) {
                # ������м�������������ŷָȥ�գ�
                $firstUserWorkstations -split "," | 
                    ForEach-Object { $_.Trim() } | 
                    Where-Object { $_ -ne "" } | 
                    ForEach-Object { $lstAllowed.Items.Add($_) | Out-Null }
            }

        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("���������ʧ�ܣ�$($_.Exception.Message)", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $restrictForm.Close()
        }
    }

    # ���ڼ���ʱִ�м��������
    $restrictForm.Add_Shown({ LoadDomainComputers })

# ---------------------- 9. �б��ƶ���ť�߼� ----------------------
    # ȫ����ӵ������б�
    $btnAddAll.Add_Click({
        $allItems = [array]$lstAllComputers.Items
        foreach ($item in $allItems) {
            if (-not $lstAllowed.Items.Contains($item)) {
                $lstAllowed.Items.Add($item) | Out-Null
            }
        }
    })

    # ѡ����ӵ������б�
    $btnAddSelected.Add_Click({
        $selectedItems = @($lstAllComputers.SelectedItems)
        foreach ($item in $selectedItems) {
            if (-not $lstAllowed.Items.Contains($item)) {
                $lstAllowed.Items.Add($item) | Out-Null
            }
        }
        # ȡ���Ҳ�ѡ�У��������飩
        $lstAllComputers.ClearSelected()
    })

    # ѡ�д������б��Ƴ�
    $btnRemoveSelected.Add_Click({
        $selectedItems = [array]$lstAllowed.SelectedItems
        foreach ($item in $selectedItems) {
            $lstAllowed.Items.Remove($item) | Out-Null
        }
        # ȡ�����ѡ��
        $lstAllowed.ClearSelected()
    })

    # ��������б�
    $btnRemoveAll.Add_Click({
        if ([System.Windows.Forms.MessageBox]::Show("ȷ��Ҫ������������¼�ļ������", "ȷ��", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq "Yes") {
            $lstAllowed.Items.Clear()
        }
    })
    

# ---------------------- 10. �����������ã������߼��� ----------------------
    $btnSave.Add_Click({
        # 1. ���������¼������б����б�ʱ���Ϊ��գ�
        $allowedComputers = ($lstAllowed.Items | ForEach-Object { $_ }) -join ","
        $isClearRestriction = [string]::IsNullOrEmpty($allowedComputers)  # �ж��Ƿ�Ϊ���������ơ�������б�

        # 2. ȷ�ϱ��棨��ʾ��Ϣ�Ż������ӡ�������ơ�˵����
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
            # ��������¼�ɹ�/ʧ�ܵ��û����飨��DisplayName��SamAccountName��
            $successUsers = @()  # ��ʽ��"��ʾ�����˺�����"
            $failUsers = @()     # ��ʽ��"��ʾ�����˺�����- ʧ��ԭ��"

            # ����ÿ��ѡ���û�����������
            foreach ($row in $selectedUsers) {
                # ��ȡ�û���DisplayName��SamAccountName����ѡ���еĵ�Ԫ���ж�ȡ��
                $displayName = $row.Cells["DisplayName"].Value
                $samAccountName = $row.Cells["SamAccountName"].Value
                # �����ֶ�Ϊ�յ������������ʾ�쳣��
                $displayName = if ([string]::IsNullOrEmpty($displayName)) { "δ������ʾ��" } else { $displayName }
                $userLabel = "$displayName��$samAccountName��"  # ͳһ���û���ʶ��ʽ

                try {
                    Invoke-Command -Session $script:remoteSession -ScriptBlock {
                        param($sam, $workstations, $isClear)
                        Import-Module ActiveDirectory -ErrorAction Stop

                        # ���������-Clear������������-Replace
                        if ($isClear) {
                            # �������ƣ����userWorkstations���ԣ������¼���м������
                            Set-ADUser -Identity $sam -Clear userWorkstations -ErrorAction Stop
                        } else {
                            # �������ƣ�����userWorkstations���ԣ�ָ�������¼�ļ������
                            Set-ADUser -Identity $sam -Replace @{userWorkstations = $workstations} -ErrorAction Stop
                        }
                    } -ArgumentList $samAccountName, $allowedComputers, $isClearRestriction -ErrorAction Stop

                    # �ɹ�����ӵ��ɹ��б�
                    $successUsers += $userLabel
                }
                catch {
                    # ʧ�ܣ���ӵ�ʧ���б���ԭ��
                    $failReason = $_.Exception.Message
                    $failUsers += "$userLabel - $failReason"
                }
            }

            # 3. ���ɽ����ʾ��������ʾ�ɹ�/ʧ���û����飩
            $resultTitle = if ($failUsers.Count -eq 0) { "�ɹ�" } else { "���" }
            $resultIcon = if ($failUsers.Count -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
            
            # ƴ�ӽ������
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

            # ��ʾ���
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
    


    # ---------------------- 11. ȡ����ť�߼� ----------------------
    $btnCancel.Add_Click({
        $restrictForm.Close()
    })

    # ---------------------- 12. ������ʾ ----------------------
    $restrictForm.Controls.Add($mainTable)
    $restrictForm.ShowDialog() | Out-Null
}
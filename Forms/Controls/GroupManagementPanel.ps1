<# 
�������壨����б�+�Ҳ������- ��ҳ�ؼ��������½ǰ汾
#>

$script:groupManagementPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:groupManagementPanel.Dock = "Fill"
$script:groupManagementPanel.ColumnCount = 2
$script:groupManagementPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$script:groupManagementPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

# ��ࣺ���б� 
$script:groupListPanel = New-Object System.Windows.Forms.GroupBox
$script:groupListPanel.Text = "���б�"
$script:groupListPanel.Dock = "Fill"
$script:groupListPanel.Padding = 10

$script:groupListTable = New-Object System.Windows.Forms.TableLayoutPanel
$script:groupListTable.Dock = "Fill"
$script:groupListTable.RowCount = 3  # ������˳�����������DataGridView����ҳ���
$script:groupListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))  # �����򣨹̶��߶ȣ�
$script:groupListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # ��DataGridView��ռ���м�ռ䣩
$script:groupListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))  # ��ҳ��壨�̶��߶ȣ�λ�ڵײ���

# �������򣨱���ԭ���߼���λ�ò��䣩
$script:groupSearchPanel = New-Object System.Windows.Forms.Panel
$script:groupSearchPanel.Dock = "Fill"

$script:labelGroupSearch = New-Object System.Windows.Forms.Label
$script:labelGroupSearch.Text = "������:"
$script:labelGroupSearch.Location = New-Object System.Drawing.Point(5, 10)
$script:labelGroupSearch.AutoSize = $true

$script:textGroupSearch = New-Object System.Windows.Forms.TextBox
$script:textGroupSearch.Location = New-Object System.Drawing.Point(80, 7)
$script:textGroupSearch.Width = 300
# �޸������¼������˺󴥷���ҳ���߼����䣩
$script:textGroupSearch.Add_TextChanged({
    $filterText = $script:textGroupSearch.Text.ToLower()
    $script:filteredGroups.Clear()

    # �����߼�����ԭFilterGroupListһ�£�
    if ([string]::IsNullOrEmpty($filterText)) {
        $script:allGroups | ForEach-Object { $script:filteredGroups.Add($_) | Out-Null }
    } else {
        $script:allGroups | Where-Object {
            $_.Name.ToLower() -like "*$filterText*" -or
            $_.SamAccountName.ToLower() -like "*$filterText*" -or
            ( (-not [string]::IsNullOrEmpty($_.Description)) -and $_.Description.ToLower() -like "*$filterText*" )
        } | ForEach-Object { $script:filteredGroups.Add($_) | Out-Null }
    }

    # ���÷�ҳ״̬����ʾ��һҳ
    $script:currentGroupPage = 1
    $script:totalGroupPages = Get-TotalPages -totalCount $script:filteredGroups.Count -pageSize $script:pageSize
    Show-GroupPage
})

$script:groupSearchPanel.Controls.Add($script:labelGroupSearch)
$script:groupSearchPanel.Controls.Add($script:textGroupSearch)
$script:groupListTable.Controls.Add($script:groupSearchPanel, 0, 0)  # ���������ڵ�1�У�����0��

# ---------------------- 2. ��DataGridView������λ�����м䣬ԭ��ҳ���λ�ã� ----------------------
$script:groupDataGridView = New-Object System.Windows.Forms.DataGridView
$script:groupDataGridView.Dock = "Fill"
$script:groupDataGridView.SelectionMode = "FullRowSelect"
$script:groupDataGridView.MultiSelect = $true
$script:groupDataGridView.ReadOnly = $true
$script:groupDataGridView.AutoGenerateColumns = $false
$script:groupDataGridView.AllowUserToAddRows = $false
$script:groupDataGridView.RowHeadersVisible = $false
$script:groupDataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 255)
$script:groupDataGridView.ColumnHeadersHeightSizeMode = "AutoSize"

# ���ж��壨���ֲ��䣩
$script:colGroupName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colGroupName.HeaderText = "������"
$script:colGroupName.DataPropertyName = "Name"
$script:colGroupName.Width = 150
$script:colGroupName.ReadOnly = $true

$script:colGroupSamAccountName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colGroupSamAccountName.HeaderText = "���˺�"
$script:colGroupSamAccountName.DataPropertyName = "SamAccountName"
$script:colGroupSamAccountName.Width = 150
$script:colGroupSamAccountName.ReadOnly = $true

$script:colGroupDescription = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colGroupDescription.HeaderText = "����"
$script:colGroupDescription.DataPropertyName = "Description"
$script:colGroupDescription.Width = 300
$script:colGroupDescription.ReadOnly = $true

$script:groupDataGridView.Columns.AddRange($script:colGroupName, $script:colGroupSamAccountName, $script:colGroupDescription)
$script:groupListTable.Controls.Add($script:groupDataGridView, 0, 1)  # ������2�У�����1����λ���������·�����ҳ�Ϸ�

# ---------------------- 3. ���ҳ��� ----------------------
$script:groupPaginationPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:groupPaginationPanel.Dock = "Fill"
$script:groupPaginationPanel.Visible = $false  # Ĭ�����أ�����>10��ʱ��ʾ��
#$script:groupPaginationPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom  # ê�����½ǣ���������ʱ����λ��
$script:groupPaginationPanel.ColumnCount = 4  # 4�У���һҳ��ť����ҳ��Ϣ����һҳ��ť
$script:groupPaginationPanel.RowCount = 1     # 1��
$script:groupPaginationPanel.Padding = New-Object System.Windows.Forms.Padding(0, 5, 10, 0)  # �Ҳ�����10px���������ߣ���������5px���Ż���ֱ���
$script:groupPaginationPanel.CellBorderStyle = "None"

# ����ʽ���ã���ť�̶���ȣ���ҳ��Ϣ����Ӧ�ı����
$script:groupPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 70)))  # 1.��һҳ
$script:groupPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))    # 2.��ҳ��Ϣ
$script:groupPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))    # 3.��ת�ؼ���������
$script:groupPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 70)))  # 4.��һҳ

# ����ʽ���̶��߶�25px��ƥ�䰴ť�߶�
$script:groupPaginationPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 25)))

# ��һҳ��ť
$script:btnGroupPrev = New-Object System.Windows.Forms.Button
$script:btnGroupPrev.Text = "��һҳ"
$script:btnGroupPrev.Width = 65
$script:btnGroupPrev.Enabled = $false
$script:btnGroupPrev.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)  # �Ҳ���5px��࣬���ҳ��Ϣ�ָ�
$script:btnGroupPrev.Add_Click({
    if ($script:currentGroupPage -gt 1) {
		$script:groupDefaultShowAll = $false  # �ر�Ĭ��ȫ��
        $script:currentGroupPage--
        Show-GroupPage
    }
})
$script:groupPaginationPanel.Controls.Add($script:btnGroupPrev, 0, 0)  # ��ӵ���1�У�����0��

# ��ҳ��Ϣ��ǩ
$script:lblGroupPageInfo = New-Object System.Windows.Forms.Label
$script:lblGroupPageInfo.Text = "�� 1 ҳ / �� 1 ҳ���ܼ� 0 ����"
$script:lblGroupPageInfo.AutoSize = $true
$script:lblGroupPageInfo.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 0)  # ������5px��ʵ�ִ�ֱ���У��Ҳ���5px������һҳ��ť�ָ�
$script:groupPaginationPanel.Controls.Add($script:lblGroupPageInfo, 1, 0)  # ��ӵ���2�У�����1��

# ��һҳ��ť
$script:btnGroupNext = New-Object System.Windows.Forms.Button
$script:btnGroupNext.Text = "��һҳ"
$script:btnGroupNext.Width = 65
$script:btnGroupNext.Enabled = $false
$script:btnGroupNext.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$script:btnGroupNext.Add_Click({
    if ($script:currentGroupPage -lt $script:totalGroupPages) {
		$script:groupDefaultShowAll = $false  # �ر�Ĭ��ȫ��
        $script:currentGroupPage++
        Show-GroupPage
    }
})
$script:groupPaginationPanel.Controls.Add($script:btnGroupNext, 2, 0)  # ��ӵ���3�У�����2��

# ����ҳ�����ӵ�groupListTable�ĵ�3�У�����2����λ��DataGridView�·�
$script:groupListTable.Controls.Add($script:groupPaginationPanel, 0, 2)

$script:groupListPanel.Controls.Add($script:groupListTable)
$script:groupManagementPanel.Controls.Add($script:groupListPanel, 0, 0)

# ---------------------- ���������ҳ��ת�ؼ� ----------------------
$script:groupJumpPanel = New-Object System.Windows.Forms.Panel
$script:groupJumpPanel.AutoSize = $true  # ����Ӧ���ݿ��
$script:groupJumpPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)

# ��ת��ǩ
$script:lblGroupJump = New-Object System.Windows.Forms.Label
$script:lblGroupJump.Text = "������"
$script:lblGroupJump.AutoSize = $true
$script:lblGroupJump.Location = New-Object System.Drawing.Point(10, 5)  # ��ֱ����

# ҳ������������������룩
$script:txtGroupJumpPage = New-Object System.Windows.Forms.TextBox
$script:txtGroupJumpPage.Width = 40  # �̶����
$script:txtGroupJumpPage.Location = New-Object System.Drawing.Point(52, 0)
$script:txtGroupJumpPage.MaxLength = 3  # �������3λ
# ֻ�����������ֺ��˸��
$script:txtGroupJumpPage.Add_KeyPress({
    if (-not ([char]::IsDigit($_.KeyChar) -or $_.KeyChar -eq [char]8)) {
        $_.Handled = $true  # ��ֹ����������
    }
})

# ��ת��ť
$script:btnGroupJump = New-Object System.Windows.Forms.Button
$script:btnGroupJump.Text = "��ת"
$script:btnGroupJump.Width = 50
$script:btnGroupJump.Location = New-Object System.Drawing.Point(100, 0)
$script:btnGroupJump.Add_Click({
    # 1. ������֤
    $jumpPage = $script:txtGroupJumpPage.Text.Trim()
    if ([string]::IsNullOrEmpty($jumpPage)) {
        [System.Windows.Forms.MessageBox]::Show("������Ҫ��ת��ҳ�룡", "��ʾ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    # �ؼ��޸�����ǰ���� $jumpPageInt ����
    $jumpPageInt = 0
    if (-not [int]::TryParse($jumpPage, [ref]$jumpPageInt)) {
        [System.Windows.Forms.MessageBox]::Show("��������Ч������ҳ�룡", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    # 2. ҳ�뷶Χ��֤
    if ($jumpPageInt -lt 1 -or $jumpPageInt -gt $script:totalGroupPages) {
        [System.Windows.Forms.MessageBox]::Show("ҳ�볬����Χ�������� 1 ~ $($script:totalGroupPages) ֮���ҳ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    # 3. ִ����ת
	$script:groupDefaultShowAll = $false
    $script:currentGroupPage = $jumpPageInt
    Show-GroupPage
})

$script:lblGroupText = New-Object System.Windows.Forms.Label
$script:lblGroupText.Text = "(��ҳ)"
$script:lblGroupText.AutoSize = $true
$script:lblGroupText.Location = New-Object System.Drawing.Point(150, 5)  # ��ֱ����

# ����ת�ؼ���ӵ������
$script:groupJumpPanel.Controls.Add($script:lblGroupJump)
$script:groupJumpPanel.Controls.Add($script:txtGroupJumpPage)
$script:groupJumpPanel.Controls.Add($script:btnGroupJump)
$script:groupJumpPanel.Controls.Add($script:lblGroupText)
# ���������ӵ���ҳ����3�У�����2��
$script:groupPaginationPanel.Controls.Add($script:groupJumpPanel, 3, 0)

# ---------------------- �Ҳࣺ�������� ----------------------
$script:groupOperationPanel = New-Object System.Windows.Forms.GroupBox
$script:groupOperationPanel.Text = "�����"
$script:groupOperationPanel.Dock = "Fill"
$script:groupOperationPanel.Padding = 15

$script:groupOpTable = New-Object System.Windows.Forms.TableLayoutPanel
$script:groupOpTable.Dock = "Fill"
$script:groupOpTable.RowCount = 5
$script:groupOpTable.ColumnCount = 2
$script:groupOpTable.Padding = 10
$script:groupOpTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$script:groupOpTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$script:groupOpTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$script:groupOpTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 80)))
$script:groupOpTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$script:groupOpTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
$script:groupOpTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

# 1. ������
$script:labelGroupName = New-Object System.Windows.Forms.Label
$script:labelGroupName.Text = "������:"
$script:labelGroupName.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:groupOpTable.Controls.Add($script:labelGroupName, 0, 0)

$script:textGroupName = New-Object System.Windows.Forms.TextBox
$script:textGroupName.Dock = "Fill"
$script:groupOpTable.Controls.Add($script:textGroupName, 1, 0)

# 2. ���˺�
$script:labelGroupSamAccount = New-Object System.Windows.Forms.Label
$script:labelGroupSamAccount.Text = "���˺�:"
$script:labelGroupSamAccount.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:groupOpTable.Controls.Add($script:labelGroupSamAccount, 0, 1)

$script:textGroupSamAccount = New-Object System.Windows.Forms.TextBox
$script:textGroupSamAccount.Dock = "Fill"
$script:groupOpTable.Controls.Add($script:textGroupSamAccount, 1, 1)

# 3. ������
$script:labelGroupDescription = New-Object System.Windows.Forms.Label
$script:labelGroupDescription.Text = "������:"
$script:labelGroupDescription.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:groupOpTable.Controls.Add($script:labelGroupDescription, 0, 2)

$script:textGroupDescription = New-Object System.Windows.Forms.TextBox
$script:textGroupDescription.Dock = "Fill"
$script:groupOpTable.Controls.Add($script:textGroupDescription, 1, 2)

# 4. �������ť
$script:groupButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$script:groupButtonsPanel.Dock = "Fill"
$script:groupButtonsPanel.Padding = New-Object System.Windows.Forms.Padding(-10, 10, 10, 10)
$script:groupButtonsPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$script:groupButtonsPanel.WrapContents = $false
$script:groupButtonsPanel.Height = 60

$script:buttonAddGroup = New-Object System.Windows.Forms.Button
$script:buttonAddGroup.Text = "�½���"
$script:buttonAddGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonAddGroup.Width = 70
$script:buttonAddGroup.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
$script:buttonAddGroup.ForeColor = [System.Drawing.Color]::White
$script:buttonAddGroup.FlatStyle = "Flat"
$script:buttonAddGroup.Add_Click({ CreateNewGroup })  # ����GroupOperations.ps1

$script:buttonAddToGroup = New-Object System.Windows.Forms.Button
$script:buttonAddToGroup.Text = "������"
$script:buttonAddToGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonAddToGroup.Width = 70
$script:buttonAddToGroup.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
$script:buttonAddToGroup.ForeColor = [System.Drawing.Color]::White
$script:buttonAddToGroup.FlatStyle = "Flat"
$script:buttonAddToGroup.Add_Click({ AddUserToGroup })  # ����GroupOperations.ps1

$script:buttonModifyGroup = New-Object System.Windows.Forms.Button
$script:buttonModifyGroup.Text = "�޸���"
$script:buttonModifyGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonModifyGroup.Width = 70
$script:buttonModifyGroup.BackColor = [System.Drawing.Color]::FromArgb(255, 140, 0)
$script:buttonModifyGroup.ForeColor = [System.Drawing.Color]::White
$script:buttonModifyGroup.FlatStyle = "Flat"
$script:buttonModifyGroup.Add_Click({ ModifyGroup })  # ����GroupOperations.ps1

$script:buttonDeleteGroup = New-Object System.Windows.Forms.Button
$script:buttonDeleteGroup.Text = "ɾ����"
$script:buttonDeleteGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonDeleteGroup.Width = 70
$script:buttonDeleteGroup.BackColor = [System.Drawing.Color]::FromArgb(220, 20, 60)
$script:buttonDeleteGroup.ForeColor = [System.Drawing.Color]::White
$script:buttonDeleteGroup.FlatStyle = "Flat"
$script:buttonDeleteGroup.Add_Click({ DeleteGroup })  # ����GroupOperations.ps1

# �����Ƴ��û���ť
$script:buttonRemoveFromGroup = New-Object System.Windows.Forms.Button
$script:buttonRemoveFromGroup.Text = "�Ƴ����˺�"
$script:buttonRemoveFromGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0) 
$script:buttonRemoveFromGroup.Width = 80 
$script:buttonRemoveFromGroup.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 70)
$script:buttonRemoveFromGroup.ForeColor = [System.Drawing.Color]::White
$script:buttonRemoveFromGroup.FlatStyle = "Flat"
$script:buttonRemoveFromGroup.Add_Click({ RemoveUserFromGroup })  # ����GroupOperations.ps1

$script:groupButtonsPanel.Controls.Add($script:buttonAddGroup)
$script:groupButtonsPanel.Controls.Add($script:buttonAddToGroup)
$script:groupButtonsPanel.Controls.Add($script:buttonModifyGroup)
$script:groupButtonsPanel.Controls.Add($script:buttonDeleteGroup)
$script:groupButtonsPanel.Controls.Add($script:buttonRemoveFromGroup)
$script:groupOpTable.Controls.Add($script:groupButtonsPanel, 1, 3)

# ��ѡ��仯�¼�������ԭ���߼���
$script:groupDataGridView.Add_SelectionChanged({
    if ($script:groupDataGridView.SelectedRows.Count -gt 0) {
        $group = $script:groupDataGridView.SelectedRows[0].DataBoundItem
        $script:textGroupName.Text = $group.Name
        $script:textGroupSamAccount.Text = $group.SamAccountName
        $script:textGroupDescription.Text = $group.Description
        $script:originalGroupSamAccount = $group.SamAccountName
    }
})

$script:groupOperationPanel.Controls.Add($script:groupOpTable)
$script:groupManagementPanel.Controls.Add($script:groupOperationPanel, 1, 0)
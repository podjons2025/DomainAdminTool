<# 
�������壨����б�+�Ҳ������ 
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
$script:groupListTable.RowCount = 2
$script:groupListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$script:groupListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

# ��������
$script:groupSearchPanel = New-Object System.Windows.Forms.Panel
$script:groupSearchPanel.Dock = "Fill"

$script:labelGroupSearch = New-Object System.Windows.Forms.Label
$script:labelGroupSearch.Text = "������:"
$script:labelGroupSearch.Location = New-Object System.Drawing.Point(5, 10)
$script:labelGroupSearch.AutoSize = $true

$script:textGroupSearch = New-Object System.Windows.Forms.TextBox
$script:textGroupSearch.Location = New-Object System.Drawing.Point(80, 7)
$script:textGroupSearch.Width = 300
$script:textGroupSearch.Add_TextChanged({ FilterGroupList $script:textGroupSearch.Text })  # ����Helpers.ps1

$script:groupSearchPanel.Controls.Add($script:labelGroupSearch)
$script:groupSearchPanel.Controls.Add($script:textGroupSearch)
$script:groupListTable.Controls.Add($script:groupSearchPanel, 0, 0)

# ��DataGridView
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

# ���ж���
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
$script:groupListTable.Controls.Add($script:groupDataGridView, 0, 1)
$script:groupListPanel.Controls.Add($script:groupListTable)
$script:groupManagementPanel.Controls.Add($script:groupListPanel, 0, 0)

#�Ҳࣺ��������
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

#�����Ƴ��û���ť
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



# ��ѡ��仯�¼�
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
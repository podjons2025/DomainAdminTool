<# 
�û�������壨����б�+�Ҳ������- ��ҳ�ؼ��������½ǰ汾
#>

$script:userManagementPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:userManagementPanel.Dock = "Fill"
$script:userManagementPanel.ColumnCount = 2
$script:userManagementPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$script:userManagementPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

# ---------------------- ��ࣺ�û��б� ----------------------
$script:userListPanel = New-Object System.Windows.Forms.GroupBox
$script:userListPanel.Text = "�˺��б�"
$script:userListPanel.Dock = "Fill"
$script:userListPanel.Padding = 10

$script:userListTable = New-Object System.Windows.Forms.TableLayoutPanel
$script:userListTable.Dock = "Fill"
$script:userListTable.RowCount = 3  # 1.������ 2.�û�DataGridView 3.��ҳ��壨�����ײ���
$script:userListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))  # ������
$script:userListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # �û�DataGridView��ռ���м䣩
$script:userListTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))  # ��ҳ��壨�̶��߶ȣ�λ�ڵײ���

# �����򣨱��ֲ��䣩
$script:searchPanel = New-Object System.Windows.Forms.Panel
$script:searchPanel.Dock = "Fill"

$script:labelSearch = New-Object System.Windows.Forms.Label
$script:labelSearch.Text = "�����˺�:"
$script:labelSearch.Location = New-Object System.Drawing.Point(5, 10)
$script:labelSearch.AutoSize = $true

$script:textSearch = New-Object System.Windows.Forms.TextBox
$script:textSearch.Location = New-Object System.Drawing.Point(80, 7)
$script:textSearch.Width = 300
# �����¼������˺󴥷���ҳ
$script:textSearch.Add_TextChanged({Update-SearchUsersResults})

$script:searchPanel.Controls.Add($script:labelSearch)
$script:searchPanel.Controls.Add($script:textSearch)
$script:userListTable.Controls.Add($script:searchPanel, 0, 0)

# ---------------------- 2. �û�DataGridView�����ֲ��䣬��������������2�У� ----------------------
$script:userDataGridView = New-Object System.Windows.Forms.DataGridView
$script:userDataGridView.Dock = "Fill"
$script:userDataGridView.SelectionMode = "FullRowSelect"
$script:userDataGridView.MultiSelect = $true
$script:userDataGridView.ReadOnly = $false
$script:userDataGridView.AutoGenerateColumns = $false
$script:userDataGridView.AllowUserToAddRows = $false
$script:userDataGridView.RowHeadersVisible = $false
$script:userDataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 255)
#$script:userDataGridView.ColumnHeadersHeightSizeMode = "AutoSize"
$script:userDataGridView.RowTemplate.Height = 20  # �и�
$script:userDataGridView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None  # �����Զ��и�
#$script:userDataGridView.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter

# �ж��壨���ֲ��䣩
$script:colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colDisplayName.HeaderText = "����"
$script:colDisplayName.DataPropertyName = "DisplayName"
$script:colDisplayName.Width = 100
$script:colDisplayName.ReadOnly = $true
$script:colDisplayName.Name = "DisplayName"

$script:colSamAccountName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colSamAccountName.HeaderText = "�˺�"
$script:colSamAccountName.DataPropertyName = "SamAccountName"
$script:colSamAccountName.Width = 110
$script:colSamAccountName.ReadOnly = $true
$script:colSamAccountName.Name = "SamAccountName" 

$script:colGroups = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colGroups.HeaderText = "������"
$script:colGroups.FillWeight = 150
$script:colGroups.ReadOnly = $true
$script:colGroups.DataPropertyName = "MemberOf"

$script:colEnabled = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$script:colEnabled.HeaderText = "�˺�״̬"
$script:colEnabled.DataPropertyName = "Enabled"
$script:colEnabled.Width = 120
$script:colEnabled.ReadOnly = $false
$script:colEnabled.FalseValue = $false
$script:colEnabled.TrueValue = $true
$script:colEnabled.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
$script:colEnabled.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$script:colEnabled.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$script:colEnabled.Name = "Enabled"

$script:colLocked = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colLocked.HeaderText = "������"
$script:colLocked.DataPropertyName = "AccountLockout"
$script:colLocked.Width = 80
$script:colLocked.ReadOnly = $true
$script:colLocked.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter

$script:colExpirationDate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colExpirationDate.HeaderText = "��������"
$script:colExpirationDate.DataPropertyName = "AccountExpirationDate"
$script:colExpirationDate.Width = 120
$script:colExpirationDate.ReadOnly = $true 
$script:colExpirationDate.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$script:colExpirationDate.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$script:colExpirationDate.Name = "AccountExpirationDate"

$script:colEmail = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colEmail.HeaderText = "����"
$script:colEmail.DataPropertyName = "EmailAddress"
$script:colEmail.Width = 180
$script:colEmail.ReadOnly = $true

$script:colPhone = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colPhone.HeaderText = "�绰"
$script:colPhone.DataPropertyName = "TelePhone"
$script:colPhone.Width = 100
$script:colPhone.ReadOnly = $true

$script:colDescription = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$script:colDescription.HeaderText = "����"
$script:colDescription.DataPropertyName = "Description"
$script:colDescription.Width = 180
$script:colDescription.ReadOnly = $true

$script:userDataGridView.Columns.AddRange($script:colDisplayName, $script:colSamAccountName, $script:colGroups, $script:colEnabled, $script:colLocked, $script:colExpirationDate, $script:colEmail, $script:colPhone, $script:colDescription)

# ��Ԫ���ʽ���¼������ֲ��䣩
$script:userDataGridView.Add_CellFormatting({
    param($sender, $e)
    if ($sender.Columns[$e.ColumnIndex].DataPropertyName -eq "AccountLockout") {
        $rawValue = $e.Value
        if ($rawValue -eq $null) { $e.Value = "��"; $e.FormattingApplied = $true; return }
        $isLocked = switch ($rawValue.GetType().Name) {
            "String" { $rawValue -eq "True" }
            "Boolean" { [bool]$rawValue }
            default { $false }
        }
        $e.Value = if ($isLocked) { "��" } else { "��" }
        $e.FormattingApplied = $true
    }
	
    # ����ʱ�����ޡ��У����ģ���ȡ����ʱ���1���ٶԱȣ�
    if ($sender.Columns[$e.ColumnIndex].DataPropertyName -eq "AccountExpirationDate") {
        $rawValue = $e.Value
        $currentDate = Get-Date -Date (Get-Date).Date  # ��ǰ���ڣ��������գ�ʱ��00:00:00��

        # 1. �������ڣ�ADδ���ù���ʱ�䣩
        if ($rawValue -eq $null -or $rawValue -is [DBNull]) {
            $e.Value = "��������"
            $e.CellStyle.ForeColor = [System.Drawing.Color]::Black
        }
        # 2. �й���ʱ�䣺��1������ж�
        elseif ($rawValue -is [DateTime]) {
            # �ؼ�����ȡ����ADʱ���1�죬���������ڲ���
            $adjustedExpiryDate = $rawValue.AddDays(-1).Date  # ��1�� + ���ʱ��
            
            # �ѹ��ڣ������������ < ��ǰ����
            if ($adjustedExpiryDate -lt $currentDate) {
                $e.Value = "�ѹ���"
                $e.CellStyle.ForeColor = [System.Drawing.Color]::Red
            }
            # δ���ڣ���ʾ����������ڣ���ADʵ�����õ�����һ�£�
            else {
                $e.Value = $adjustedExpiryDate.ToString("yyyy-MM-dd")
                $e.CellStyle.ForeColor = [System.Drawing.Color]::Black
            }
        }
        # 3. �쳣ֵ����
        else {
            $e.Value = "δ֪"
            $e.CellStyle.ForeColor = [System.Drawing.Color]::Gray
        }

        $e.FormattingApplied = $true 
    }	
})

# �˺�״̬�л����¼������ֲ��䣩
$script:userDataGridView.Add_CellPainting({
    param($sender, $e)
    if ($e.ColumnIndex -ge 0 -and $sender.Columns[$e.ColumnIndex].Name -eq "Enabled" -and $e.RowIndex -ge 0) {
        $e.PaintBackground($e.CellBounds, $true)
        $e.PaintContent($e.CellBounds)
        $cellValue = $sender.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value
        $isEnabled = if ($cellValue -eq $null) { $false } else { [bool]$cellValue }
        $statusText = if ($isEnabled) { "����" } else { "����" }
        $statusColor = if ($isEnabled) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red }
        
        $textFormat = [System.Drawing.StringFormat]::new()
        $textFormat.Alignment = [System.Drawing.StringAlignment]::Near
        $textFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
        $textRect = [System.Drawing.RectangleF]::new($e.CellBounds.X + 25, $e.CellBounds.Y, $e.CellBounds.Width - 30, $e.CellBounds.Height)
        $font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $e.Graphics.DrawString($statusText, $font, [System.Drawing.SolidBrush]$statusColor, $textRect, $textFormat)
        
        $textFormat.Dispose()
        $font.Dispose()
        $e.Handled = $true
    }
})

# �˺�״̬�л��¼�������/���ã������ֲ��䣩
$script:userDataGridView.Add_CellContentClick({
    if ($_.ColumnIndex -eq $script:colEnabled.Index -and $_.RowIndex -ge 0) {
        ToggleUserEnabled $_.RowIndex  # ����UserOperations.ps1
    }
})

# �û�ѡ��仯�¼������ֲ��䣩
$script:userDataGridView.Add_SelectionChanged({
    if ($script:userDataGridView.SelectedRows.Count -gt 0) {
        $user = $script:userDataGridView.SelectedRows[0].DataBoundItem
        $script:textCnName.Text = $user.DisplayName 
        $script:textPinyin.Text = $user.SamAccountName
        $script:textEmail.Text = $user.EmailAddress
		$script:textPhone.Text = $user.TelePhone
        $script:textDescription.Text = $user.Description
        $script:textNewPassword.Text = ""
        $script:textConfirmPassword.Text = ""
    }
})

$script:userListTable.Controls.Add($script:userDataGridView, 0, 1)  # ������2�У�����1��

# ---------------------- 3. ��ҳ��� ----------------------
# �ع�ΪTableLayoutPanel��ʵ�����½Ƕ���
$script:userPaginationPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:userPaginationPanel.Dock = "Fill"  # ��丸������userListTable�ĵ�3�У�
$script:userPaginationPanel.Visible = $false  # Ĭ�����أ�����>6��ʱ��ʾ��
$script:userPaginationPanel.ColumnCount = 4 
$script:userPaginationPanel.RowCount = 1     # 1��
#$script:userPaginationPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom  # ���½�ê��
$script:userPaginationPanel.Padding = New-Object System.Windows.Forms.Padding(0, 5, 10, 0)  # �Ҳ����գ���������
$script:userPaginationPanel.CellBorderStyle = "None"


# ����ʽ����ť�̶���ȣ���ҳ��Ϣ����Ӧ
$script:userPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 70)))  # 1.��һҳ
$script:userPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))    # 2.��ҳ��Ϣ
$script:userPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))    # 3.��ת�ؼ���������
$script:userPaginationPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 70)))  # 4.��һҳ
$script:userPaginationPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 25)))

# ��һҳ��ť���Ƴ��̶�Location����TableLayoutPanel����λ�ã�
$script:btnUserPrev = New-Object System.Windows.Forms.Button
$script:btnUserPrev.Text = "��һҳ"
$script:btnUserPrev.Width = 65
$script:btnUserPrev.Enabled = $false
$script:btnUserPrev.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)  # ���ҳ��Ϣ�����
$script:btnUserPrev.Add_Click({
    if ($script:currentUserPage -gt 1) {
		$script:defaultShowAll = $false  # �ر�Ĭ��ȫ�ԣ������ҳģʽ
        $script:currentUserPage--
        Show-UserPage
    }
})
$script:userPaginationPanel.Controls.Add($script:btnUserPrev, 0, 0)  # ��1�У�����0��
$script:userPaginationPanel.SetColumnSpan($script:btnUserPrev, 1)
$script:userPaginationPanel.SetRowSpan($script:btnUserPrev, 1)

# ��ҳ��Ϣ��ǩ
$script:lblUserPageInfo = New-Object System.Windows.Forms.Label
$script:lblUserPageInfo.Text = "�� 1 ҳ / �� 1 ҳ���ܼ� 0 ����"
$script:lblUserPageInfo.AutoSize = $true
$script:lblUserPageInfo.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 0)  # ��ֱ���ж���
$script:userPaginationPanel.Controls.Add($script:lblUserPageInfo, 1, 0)  # ��2�У�����1��
$script:userPaginationPanel.SetColumnSpan($script:lblUserPageInfo, 1)
$script:userPaginationPanel.SetRowSpan($script:lblUserPageInfo, 1)

# ��һҳ��ť
$script:btnUserNext = New-Object System.Windows.Forms.Button
$script:btnUserNext.Text = "��һҳ"
$script:btnUserNext.Width = 65
$script:btnUserNext.Enabled = $false
$script:btnUserNext.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$script:btnUserNext.Add_Click({
    if ($script:currentUserPage -lt $script:totalUserPages) {
		$script:defaultShowAll = $false  # �ر�Ĭ��ȫ��
        $script:currentUserPage++
        Show-UserPage
    }
})
$script:userPaginationPanel.Controls.Add($script:btnUserNext, 2, 0)  # ��3�У�����2��
$script:userPaginationPanel.SetColumnSpan($script:btnUserNext, 1)
$script:userPaginationPanel.SetRowSpan($script:btnUserNext, 1)

# ����ҳ�����ӵ�userListTable�ĵ�3�У�����2��
$script:userListTable.Controls.Add($script:userPaginationPanel, 0, 2)

$script:userListPanel.Controls.Add($script:userListTable)
$script:userManagementPanel.Controls.Add($script:userListPanel, 0, 0)


# ---------------------- �������û���ҳ��ת�ؼ� ----------------------
$script:userJumpPanel = New-Object System.Windows.Forms.Panel
$script:userJumpPanel.AutoSize = $true  # ����Ӧ���ݿ��
$script:userJumpPanel.Margin = New-Object System.Windows.Forms.Padding(5, 0, 10, 0)

# ��ת��ǩ
$script:lblUserJump = New-Object System.Windows.Forms.Label
$script:lblUserJump.Text = "������"
$script:lblUserJump.AutoSize = $true
$script:lblUserJump.Location = New-Object System.Drawing.Point(10, 5)  # ��ֱ����

# ҳ������������������룩
$script:txtUserJumpPage = New-Object System.Windows.Forms.TextBox
$script:txtUserJumpPage.Width = 40  # �̶���ȣ�����Ƶ���仯
$script:txtUserJumpPage.Location = New-Object System.Drawing.Point(52, 0)
$script:txtUserJumpPage.MaxLength = 3  # �����������3λ�����999ҳ��
# ֻ�����������ֺ��˸��
$script:txtUserJumpPage.Add_KeyPress({
    if (-not ([char]::IsDigit($_.KeyChar) -or $_.KeyChar -eq [char]8)) {
        $_.Handled = $true  # ��ֹ����������
    }
})

# ��ת��ť
$script:btnUserJump = New-Object System.Windows.Forms.Button
$script:btnUserJump.Text = "��ת"
$script:btnUserJump.Width = 50
$script:btnUserJump.Location = New-Object System.Drawing.Point(100, 0)
$script:btnUserJump.Add_Click({
    # 1. ������֤
    $jumpPage = $script:txtUserJumpPage.Text.Trim()
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
    # 2. ҳ�뷶Χ��֤��1 ~ ��ҳ����
    if ($jumpPageInt -lt 1 -or $jumpPageInt -gt $script:totalUserPages) {
        [System.Windows.Forms.MessageBox]::Show("ҳ�볬����Χ�������� 1 ~ $($script:totalUserPages) ֮���ҳ��", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    # 3. ִ����ת
	$script:defaultShowAll = $false
    $script:currentUserPage = $jumpPageInt
    Show-UserPage
})

$script:lblUserText = New-Object System.Windows.Forms.Label
$script:lblUserText.Text = "(��ҳ)"
$script:lblUserText.AutoSize = $true
$script:lblUserText.Location = New-Object System.Drawing.Point(150, 5)  # ��ֱ����


# ����ת�ؼ���ӵ������
$script:userJumpPanel.Controls.Add($script:lblUserJump)
$script:userJumpPanel.Controls.Add($script:txtUserJumpPage)
$script:userJumpPanel.Controls.Add($script:btnUserJump)
$script:userJumpPanel.Controls.Add($script:lblUserText)
# ���������ӵ���ҳ����3�У�����2��
$script:userPaginationPanel.Controls.Add($script:userJumpPanel, 3, 0)





# ---------------------- �Ҳࣺ�û�������� ----------------------
$script:userOperationPanel = New-Object System.Windows.Forms.GroupBox
$script:userOperationPanel.Text = "�˺Ų���"
$script:userOperationPanel.Dock = "Fill"
$script:userOperationPanel.Padding = New-Object System.Windows.Forms.Padding(20, 10, 10, 10)  # ���ӿ��ڱ߾࣬�����ǩӵ��

$script:operationTable = New-Object System.Windows.Forms.TableLayoutPanel
$script:operationTable.Dock = "Fill"
$script:operationTable.RowCount = 7  # 7�нṹ����������¼+ǰ׺�����䡢���������ڡ������롢ȷ�����룩
$script:operationTable.ColumnCount = 2  # 2�У���ǩ�� + �ؼ���
$script:operationTable.Padding = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
$script:operationTable.CellBorderStyle = "None"

# ͳһ�и�Ϊ35px����֤��ֱ������
for ($i=0; $i -lt $script:operationTable.RowCount; $i++) {
    $script:operationTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
}
$script:operationTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))  # ��ǩ�й̶����
$script:operationTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # �ؼ���ռ��ʣ����

# ---------- 1. ���� ----------
$script:labelCnName = New-Object System.Windows.Forms.Label
$script:labelCnName.Text = "����:"
$script:labelCnName.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # ��ǩ�����Ҿ���
$script:labelCnName.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)  # �Ҳ����Ӽ�࣬��������Э��
$script:operationTable.Controls.Add($script:labelCnName, 0, 0)

$script:textCnName = New-Object System.Windows.Forms.TextBox
$script:textCnName.Dock = "Fill"
$script:textCnName.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)
$script:textCnName.Add_TextChanged({ ConvertToPinyin })  # ����PinyinConverter.ps1
$script:operationTable.Controls.Add($script:textCnName, 1, 0)

# ---------- 2. ��¼�˺� + �˺�ǰ׺ ----------
$script:labelPinyin = New-Object System.Windows.Forms.Label
$script:labelPinyin.Text = "��¼�˺�:"
$script:labelPinyin.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # �롰��������ǩ����
$script:labelPinyin.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelPinyin, 0, 1)

$script:accountSubPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:accountSubPanel.Dock = "Fill"
$script:accountSubPanel.ColumnCount = 3  # 3�У���¼�����ǰ׺��ǩ��ǰ׺�����
$script:accountSubPanel.RowCount = 1
$script:accountSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
$script:accountSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 80)))
$script:accountSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$script:accountSubPanel.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)

$script:textPinyin = New-Object System.Windows.Forms.TextBox
$script:textPinyin.Dock = "Fill"
$script:textPinyin.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)  # ��ǰ׺��ǩ�����
$script:accountSubPanel.Controls.Add($script:textPinyin, 0, 0)

$script:labelPrefix = New-Object System.Windows.Forms.Label
$script:labelPrefix.Text = "�˺�ǰ׺:"
$script:labelPrefix.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:labelPrefix.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)
$script:accountSubPanel.Controls.Add($script:labelPrefix, 1, 0)

$script:textPrefix = New-Object System.Windows.Forms.TextBox
$script:textPrefix.Text = "IBM_"
$script:textPrefix.Dock = "Fill"
$script:textPrefix.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$script:textPrefix.Add_TextChanged({ ConvertToPinyin })  # ����PinyinConverter.ps1
$script:accountSubPanel.Controls.Add($script:textPrefix, 2, 0)

$script:operationTable.Controls.Add($script:accountSubPanel, 1, 1)

# ---------- 3. ���� + �绰 ----------
$script:labelEmail = New-Object System.Windows.Forms.Label
$script:labelEmail.Text = "����:"
$script:labelEmail.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight 
$script:labelEmail.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelEmail, 0, 2)

# �Ҳഴ�������������+�绰��ǩ+�绰����򡱵������
$script:contactSubPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:contactSubPanel.Dock = "Fill"
$script:contactSubPanel.ColumnCount = 4
$script:contactSubPanel.RowCount = 1
$script:contactSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
$script:contactSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 80)))
$script:contactSubPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$script:contactSubPanel.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)

# 1. ���������
$script:textEmail = New-Object System.Windows.Forms.TextBox
$script:textEmail.Dock = "Fill"
$script:textEmail.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0) 
$script:contactSubPanel.Controls.Add($script:textEmail, 0, 0)

# 2. �绰��ǩ
$script:labelPhone = New-Object System.Windows.Forms.Label
$script:labelPhone.Text = "��ϵ�绰:"
$script:labelPhone.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight 
$script:labelPhone.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0) 
$script:contactSubPanel.Controls.Add($script:labelPhone, 1, 0)

# 3. �绰�����
$script:textPhone = New-Object System.Windows.Forms.TextBox
$script:textPhone.Dock = "Fill"
$script:textPhone.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$script:contactSubPanel.Controls.Add($script:textPhone, 2, 0)

# ��ֹ����������ַ���
$script:textPhone.Add_KeyPress({
    # ֱ�������ַ����͵����ַ�Χ��'0'��'9'��char���ͣ��������˸����ASCII 8��
    $allowedKeys = @([char]8) + ([char]'0'..[char]'9')  # ��[char]��ʽָ���ַ�����
    
    # ��鵱ǰ���µļ��Ƿ��������б���
    if ($allowedKeys -notcontains $_.KeyChar) {
        $_.Handled = $true  # ��ֹ�������ַ�����
		[System.Windows.Forms.MessageBox]::Show("��������ȷ�ĵ绰���룡����", "����", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# ���������ӵ������
$script:operationTable.Controls.Add($script:contactSubPanel, 1, 2)

# ---------- 4. ���� ----------
$script:labelDescription = New-Object System.Windows.Forms.Label
$script:labelDescription.Text = "����:"
$script:labelDescription.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:labelDescription.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelDescription, 0, 3)

$script:textDescription = New-Object System.Windows.Forms.TextBox
$script:textDescription.Dock = "Fill"
$script:textDescription.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)
$script:operationTable.Controls.Add($script:textDescription, 1, 3)

# ---------- 5. �������� ----------
$script:labelExpiry = New-Object System.Windows.Forms.Label
$script:labelExpiry.Text = "��������:"
$script:labelExpiry.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:labelExpiry.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelExpiry, 0, 4)

$script:expiryPanel = New-Object System.Windows.Forms.Panel
$script:expiryPanel.Dock = "Fill"
$script:expiryPanel.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)

# ����ѡ����
$script:dateExpiry = New-Object System.Windows.Forms.DateTimePicker
$script:dateExpiry.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$script:dateExpiry.Location = New-Object System.Drawing.Point(0, 5)
$script:dateExpiry.Width = 120
$script:dateExpiry.Height = 22
$script:dateExpiry.Enabled = $false

# �������ڸ�ѡ��
$script:chkNeverExpire = New-Object System.Windows.Forms.CheckBox
$script:chkNeverExpire.Text = "��������"
$script:chkNeverExpire.Location = New-Object System.Drawing.Point(135, 5)
$script:chkNeverExpire.Width = 80
$script:chkNeverExpire.Checked = $true
$script:chkNeverExpire.Add_CheckedChanged({ $script:dateExpiry.Enabled = -not $this.Checked })

# ��ʼ�����ǩ��λ�ڡ�¼�����롱��ť��ࣩ
$script:labelPassword = New-Object System.Windows.Forms.Label
$script:labelPassword.Text = "(���룺Password@001)"
$script:labelPassword.Height = 25
$script:labelPassword.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$script:labelPassword.Location = New-Object System.Drawing.Point(213, 8)  # ����λ�õ���ť���
$script:labelPassword.AutoSize = $true  # �Զ���Ӧ�ı����


# ��ʼ���밴ť�����ƣ����ʼ�����ǩ���룩
$script:buttonCreatePassword = New-Object System.Windows.Forms.Button
$script:buttonCreatePassword.Text = "��ʼ����"
$script:buttonCreatePassword.Location = New-Object System.Drawing.Point(340, 5)  # λ�ڳ�ʼ�����ǩ�Ҳ�
$script:buttonCreatePassword.Width = 80
$script:buttonCreatePassword.Height = 25
$script:buttonCreatePassword.BackColor = [System.Drawing.Color]::FromArgb(100, 150, 250)
$script:buttonCreatePassword.ForeColor = [System.Drawing.Color]::White
$script:buttonCreatePassword.FlatStyle = "Flat"
# ����¼��������ʼ����
$script:buttonCreatePassword.Add_Click({ 
    $script:textNewPassword.Text = "Password@001"
    $script:textConfirmPassword.Text = "Password@001"
})

# �����пؼ���ӵ����
$script:expiryPanel.Controls.Add($script:dateExpiry)
$script:expiryPanel.Controls.Add($script:chkNeverExpire)
$script:expiryPanel.Controls.Add($script:labelPassword)  # ������ʼ�����ǩ
$script:expiryPanel.Controls.Add($script:buttonCreatePassword)

$script:operationTable.Controls.Add($script:expiryPanel, 1, 4)

# �����루�Զ������ʼ���룩
$script:labelNewPassword = New-Object System.Windows.Forms.Label
$script:labelNewPassword.Text = "������:"
$script:labelNewPassword.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:labelNewPassword.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelNewPassword, 0, 5)

$script:textNewPassword = New-Object System.Windows.Forms.TextBox
$script:textNewPassword.PasswordChar = '*'
$script:textNewPassword.Dock = "Fill"
$script:textNewPassword.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 5)
$script:textNewPassword.Text = ""
$script:operationTable.Controls.Add($script:textNewPassword, 1, 5)

# ȷ�����루�Զ������ʼ���룬���ӵײ����գ�
$script:labelConfirmPassword = New-Object System.Windows.Forms.Label
$script:labelConfirmPassword.Text = "ȷ������:"
$script:labelConfirmPassword.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:labelConfirmPassword.Margin = New-Object System.Windows.Forms.Padding(5, 5, 10, 5)
$script:operationTable.Controls.Add($script:labelConfirmPassword, 0, 6)

$script:textConfirmPassword = New-Object System.Windows.Forms.TextBox
$script:textConfirmPassword.PasswordChar = '*'
$script:textConfirmPassword.Dock = "Fill"
$script:textConfirmPassword.Margin = New-Object System.Windows.Forms.Padding(0, 5, 5, 10)
$script:textConfirmPassword.Text = ""
$script:operationTable.Controls.Add($script:textConfirmPassword, 1, 6)

# ͳһ�����ı�����С�߶�
foreach ($tb in $script:operationTable.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] }) {
    $tb.MinimumSize = New-Object System.Drawing.Size(0, 22)
}

$script:userOperationPanel.Controls.Add($script:operationTable)
$script:userManagementPanel.Controls.Add($script:userOperationPanel, 1, 0)

# �м������ť�����ø�����壬���ֲ��䣩
$script:middleButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$script:middleButtonPanel.Dock = "Fill"
$script:middleButtonPanel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
$script:middleButtonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
$script:middleButtonPanel.WrapContents = $false
$script:middleButtonPanel.Height = 60

$script:buttonCreate = New-Object System.Windows.Forms.Button
$script:buttonCreate.Text = "�½��˺�"
$script:buttonCreate.Width = 80
$script:buttonCreate.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonCreate.BackColor = [System.Drawing.Color]::ForestGreen
$script:buttonCreate.ForeColor = [System.Drawing.Color]::White
$script:buttonCreate.FlatStyle = "Flat"
$script:buttonCreate.Add_Click({ CreateNewUser })  # ����UserOperations.ps1

$script:buttonChangePassword = New-Object System.Windows.Forms.Button
$script:buttonChangePassword.Text = "�޸�����"
$script:buttonChangePassword.Width = 80
$script:buttonChangePassword.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonChangePassword.BackColor = [System.Drawing.Color]::Orange
$script:buttonChangePassword.ForeColor = [System.Drawing.Color]::White
$script:buttonChangePassword.FlatStyle = "Flat"
$script:buttonChangePassword.Add_Click({ ChangeUserPassword })  # ����UserOperations.ps1

$script:buttonModifyUser = New-Object System.Windows.Forms.Button
$script:buttonModifyUser.Text = "�޸���Ϣ"
$script:buttonModifyUser.Width = 80
$script:buttonModifyUser.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonModifyUser.BackColor = [System.Drawing.Color]::DarkCyan
$script:buttonModifyUser.ForeColor = [System.Drawing.Color]::White
$script:buttonModifyUser.FlatStyle = "Flat"
$script:buttonModifyUser.Add_Click({ ModifyUserAccount })  # ����UserOperations.ps1

$script:buttonUnlock = New-Object System.Windows.Forms.Button
$script:buttonUnlock.Text = "�����˺�"
$script:buttonUnlock.Width = 80
$script:buttonUnlock.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonUnlock.BackColor = [System.Drawing.Color]::DarkOrchid
$script:buttonUnlock.ForeColor = [System.Drawing.Color]::White
$script:buttonUnlock.FlatStyle = "Flat"
$script:buttonUnlock.Add_Click({ UnlockUserAccount })  # ����UserOperations.ps1

$script:buttonRefresh = New-Object System.Windows.Forms.Button
$script:buttonRefresh.Text = "ˢ���б�"
$script:buttonRefresh.Width = 80
$script:buttonRefresh.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonRefresh.BackColor = [System.Drawing.Color]::SteelBlue
$script:buttonRefresh.ForeColor = [System.Drawing.Color]::White
$script:buttonRefresh.FlatStyle = "Flat"
$script:buttonRefresh.Add_Click({ LoadUserList; LoadGroupList})  # ����User/GroupOperations.ps1

$script:buttonRename = New-Object System.Windows.Forms.Button
$script:buttonRename.Text = "�������˺�"
$script:buttonRename.Width = 80
$script:buttonRename.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonRename.BackColor = [System.Drawing.Color]::MediumPurple
$script:buttonRename.ForeColor = [System.Drawing.Color]::White
$script:buttonRename.FlatStyle = "Flat"
$script:buttonRename.Add_Click({ RenameUserAccount })  # ����UserOperations.ps1

$script:buttonDelete = New-Object System.Windows.Forms.Button
$script:buttonDelete.Text = "ɾ���˺�"
$script:buttonDelete.Width = 80
$script:buttonDelete.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonDelete.BackColor = [System.Drawing.Color]::Crimson
$script:buttonDelete.ForeColor = [System.Drawing.Color]::White
$script:buttonDelete.FlatStyle = "Flat"
$script:buttonDelete.Add_Click({ DeleteUserAccount })  # ����UserOperations.ps1

# ��������CSV��ť
$script:buttonImportCSV = New-Object System.Windows.Forms.Button
$script:buttonImportCSV.Text = "����CSV��������"
$script:buttonImportCSV.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonImportCSV.Width = 110
$script:buttonImportCSV.BackColor = [System.Drawing.Color]::FromArgb(30, 100, 120)
$script:buttonImportCSV.ForeColor = [System.Drawing.Color]::White
$script:buttonImportCSV.FlatStyle = "Flat"
$script:buttonImportCSV.Add_Click({ImportCSVAndCreateUsers})  # ����importExportUsers.ps1

# ����CSV��ť
$script:buttonExportCSV = New-Object System.Windows.Forms.Button
$script:buttonExportCSV.Text = "����CSV"
$script:buttonExportCSV.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$script:buttonExportCSV.Width = 90
$script:buttonExportCSV.BackColor = [System.Drawing.Color]::FromArgb(150, 120, 80)
$script:buttonExportCSV.ForeColor = [System.Drawing.Color]::White
$script:buttonExportCSV.FlatStyle = "Flat"
$script:buttonExportCSV.Add_Click({ExportCSVUsers})   # ����importExportUsers.ps1

$script:middleButtonPanel.Controls.AddRange(@($script:buttonDelete, `
$script:buttonChangePassword, $script:buttonModifyUser, $script:buttonUnlock,`
$script:buttonRefresh, $script:buttonRename, $script:buttonCreate, `
$script:buttonExportCSV, $script:buttonImportCSV
))
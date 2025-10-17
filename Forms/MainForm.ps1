<# 
�����嶨�� 
#>

$script:mainForm = New-Object System.Windows.Forms.Form
$script:mainForm.Text = "����˺Ź�����"
$script:mainForm.Size = New-Object System.Drawing.Size(1200, 900)
$script:mainForm.StartPosition = "CenterScreen"
#$script:mainForm.FormBorderStyle = "Fixed3D"
#$script:mainForm.MaximizeBox = $false
$script:mainForm.FormBorderStyle = "Sizable"
$script:mainForm.MaximizeBox = $true
$script:mainForm.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)

# �����ڴ�С�仯�¼���ǿ��ˢ�����и�����+�ӳ����䣩
$script:mainForm.Add_SizeChanged({
    # 1. ǿ��ˢ������Ƕ�׸���壨ȷ���ӿؼ��ߴ�ͬ�����£�
    $script:mainPanel.PerformLayout()          # ����������
    $script:userManagementPanel.PerformLayout()# �û��������
    $script:groupManagementPanel.PerformLayout()# ��������
    $script:userListPanel.PerformLayout()      # �û��б������
    $script:groupListPanel.PerformLayout()     # ���б������
	$script:ouButtonPanel.PerformLayout()      # ˢ��OU��ť��岼��

    # 2. ���/��ԭ���ӳ�50ms���ȴ�ϵͳ��ɲ��֣�����С��������
    if ($script:mainForm.WindowState -in [System.Windows.Forms.FormWindowState]::Maximized, [System.Windows.Forms.FormWindowState]::Normal) {
        Start-Sleep -Milliseconds 50  # ��ϵͳ�㹻ʱ����¿ؼ��ߴ�
        Update-DynamicUserPageSize
        Update-DynamicGroupPageSize
    }
})

# DataGridView���������������״ζ�̬�����������ʼֵΪ0��
$script:userDataGridView.PerformLayout()  # ȷ��DGV�����ѳ�ʼ��
Update-DynamicUserPageSize  # ��ǰ���㶯̬����

$script:userDataGridView.Add_SizeChanged({
    Start-Sleep -Milliseconds 50
    Update-DynamicUserPageSize
})

# DataGridView���������������״ζ�̬�����������ʼֵΪ0��
$script:groupDataGridView.PerformLayout()  # ȷ��DGV�����ѳ�ʼ��
Update-DynamicGroupPageSize  # ��ǰ���㶯̬����

$script:groupDataGridView.Add_SizeChanged({
    Start-Sleep -Milliseconds 50
    Update-DynamicGroupPageSize
})

# ��������壨5�нṹ��
$script:mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:mainPanel.Dock = "Fill"
$script:mainPanel.RowCount = 6 
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 160)))  # ��������������߶�
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45)))  # �ϲ㰴ť
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 45)))   # �û�����
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45)))  # �м䰴ť
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 45)))   # �����
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))  # ״̬��ʾ

# 1. ���������壨��0�У�
$script:mainPanel.Controls.Add($script:connectionPanel, 0, 0)

# 2. OU������ť��壨��1�У�����һ�У��Ƴ��������GroupBox��
$script:mainPanel.Controls.Add($script:ouButtonPanel, 0, 1)

# 3. �û�������壨��2�У�ԭ����������1λ��
$script:mainPanel.Controls.Add($script:userManagementPanel, 0, 2)

# 4. �м������ť����3�У�ԭ����������1λ��
$script:mainPanel.Controls.Add($script:middleButtonPanel, 0, 3)

# 5. �������壨��4�У�ԭ����������1λ��
$script:mainPanel.Controls.Add($script:groupManagementPanel, 0, 4)

# 6. ״̬������5�У�ԭ����������1λ��
$script:mainPanel.Controls.Add($script:statusOutputLabel, 0, 5)

# ���������ӵ�������
$script:mainForm.Controls.Add($script:mainPanel)

# �����ڼ�������¼���������������ʾ���ټ����ʼ������
$script:mainForm.Add_Load({
    Start-Sleep -Milliseconds 100  # �ȴ�������ȫ��Ⱦ
    Update-DynamicUserPageSize
    Update-DynamicGroupPageSize
    $script:ouButtonPanel.PerformLayout()  # ˢ��OU��ť���
    # ǿ��ˢ��DataGridView��ȷ��������״̬����
    $script:userDataGridView.Refresh()
    $script:groupDataGridView.Refresh()
})

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

# ��������壨5�нṹ��
$script:mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$script:mainPanel.Dock = "Fill"
$script:mainPanel.RowCount = 5
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 180)))  # ��������������߶�
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 45)))   # �û�����
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45)))  # �м䰴ť
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 45)))   # �����
$script:mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))  # ״̬��ʾ

# 1. �������������
$script:mainPanel.Controls.Add($script:connectionPanel, 0, 0)

# 2. ����û��������
$script:mainPanel.Controls.Add($script:userManagementPanel, 0, 1)

# 3. ����м������ť
$script:mainPanel.Controls.Add($script:middleButtonPanel, 0, 2)

# 4. �����������
$script:mainPanel.Controls.Add($script:groupManagementPanel, 0, 3)

# 5. ���״̬��
$script:mainPanel.Controls.Add($script:statusOutputLabel, 0, 4)

# ���������ӵ�������
$script:mainForm.Controls.Add($script:mainPanel)

<# 
�ײ�״̬��ʾ�� 
#>

$script:statusOutputLabel = New-Object System.Windows.Forms.Label
$script:statusOutputLabel.Dock = "Fill"
$script:statusOutputLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$script:statusOutputLabel.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
$script:statusOutputLabel.Text = "δ���ӵ���� | �û���: 0 | ����: 0"

# ����״̬����ȫ�ֿ��ã�
function UpdateStatusBar {
    $script:statusOutputLabel.Text = "$($script:connectionStatus) | �Ѽ��� $($script:userCountStatus) ���û� | �Ѽ��� $($script:groupCountStatus) ����"
}
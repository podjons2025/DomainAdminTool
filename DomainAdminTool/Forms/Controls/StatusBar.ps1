<# 
底部状态显示栏 
#>

$script:statusOutputLabel = New-Object System.Windows.Forms.Label
$script:statusOutputLabel.Dock = "Fill"
$script:statusOutputLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$script:statusOutputLabel.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
$script:statusOutputLabel.Text = "未连接到域控 | 用户数: 0 | 组数: 0"

# 更新状态栏（全局可用）
function UpdateStatusBar {
    $script:statusOutputLabel.Text = "$($script:connectionStatus) | 已加载 $($script:userCountStatus) 个用户 | 已加载 $($script:groupCountStatus) 个组"
}
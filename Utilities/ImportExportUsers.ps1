<#
.SYNOPSIS
AD用户批量导入出函数
#>


function ImportCSVAndCreateUsers {	
    if (-not $script:domainContext) {
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }	

    # 打开文件选择对话框
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "CSV文件 (*.csv)|*.csv|所有文件 (*.*)|*.*"
    $fileDialog.Title = "选择包含用户信息的CSV文件"
    
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvPath = $fileDialog.FileName
        $script:connectionStatus = "正在处理CSV文件: $([System.IO.Path]::GetFileName($csvPath))"
        $script:mainForm.Refresh()

        try {
            # 检查CSV文件是否存在
            if (-not (Test-Path -Path $csvPath -PathType Leaf)) {
                [System.Windows.Forms.MessageBox]::Show("CSV文件不存在: $csvPath", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            # 读取CSV文件
            try {
                $users = Import-Csv -Path $csvPath -Encoding Default -ErrorAction Stop
                $userCount = $users | Measure-Object | Select-Object -ExpandProperty Count
                $script:connectionStatus = "成功读取CSV文件，共发现 $userCount 个用户记录"
                $script:mainForm.Refresh()
            } catch {
                [System.Windows.Forms.MessageBox]::Show("读取CSV文件失败: $_", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            # 显示确认对话框
            $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                "发现 $userCount 个账号记录，是否确定导入并创建这些用户？",
                "确认导入",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
                $script:connectionStatus = "用户取消了导入操作"
                return
            }

            # 将用户数据序列化为字符串，以便在远程会话中使用
            $usersJson = $users | ConvertTo-Json
            
            $script:connectionStatus = "正在远程创建用户..."
            $script:mainForm.Refresh()

            # 执行远程操作
            $result = Invoke-Command -Session $script:remoteSession -ScriptBlock {
                param($usersJson, $NameOU)
                
                # 初始化结果对象，添加详细信息属性
                $result = [PSCustomObject]@{
                    TotalUsers = 0
                    CreatedUsers = 0
                    SkippedUsers = 0
                    ExistingUsers = @()  # 存储已存在的用户信息
                    TotalGroups = 0
                    CreatedGroups = 0
                    CreatedGroupsDetails = @()  # 存储创建成功的组信息
                    ErrorLogs = @()
                    CreatedUsersDetails = @()  # 存储创建成功的用户详情
                }

                # 反序列化用户数据
                try {
                    $users = $usersJson | ConvertFrom-Json
                    $result.TotalUsers = $users.Count
                }
                catch {
                    $result.ErrorLogs += "解析用户数据失败: $_"
                    return $result
                }

                # 检查AD模块
                if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
                    $result.ErrorLogs += "Active Directory模块未安装。请安装RSAT工具中的AD模块。"
                    return $result
                }
                try {
                    Import-Module ActiveDirectory -ErrorAction Stop
                }
                catch {
                    $result.ErrorLogs += "导入Active Directory模块失败: $_"
                    return $result
                }

                # -------------------------- 复姓列表（保留原功能） --------------------------
                $doubleSurnames = @(
                    "欧阳", "太史", "端木", "上官", "司马", "东方", "独孤", "南宫",
                    "万俟", "闻人", "夏侯", "诸葛", "尉迟", "公羊", "赫连", "澹台",
                    "皇甫", "宗政", "濮阳", "公冶", "太叔", "申屠", "公孙", "慕容",
                    "钟离", "长孙", "司徒", "鲜于", "司空", "亓官", "司寇", "仉督",
                    "子车", "颛孙", "端木", "巫马", "公西", "漆雕", "乐正", "壤驷",
                    "公良", "拓跋", "夹谷", "宰父", "谷梁", "段干", "百里", "呼延",
                    "东郭", "南门", "羊舌", "微生", "左丘", "东门", "西门", "第五",
                    "言福", "刘付", "相里", "子书", "即墨", "达奚", "褚师", "况后",
                    "梁丘", "东宫", "仲长", "屈突", "尔朱", "纳兰", "慕容", "尉迟",
                    "可频", "纥豆陵", "宿勤", "阿跌", "斛律", "叱吕", "贺若", "秃发",
                    "乞伏", "厍狄", "乌古论", "古里", "夹谷", "蒲察", "女奚烈", "兀颜",
                    "斯陈", "孙伯", "归海", "后氏", "有氏", "琴氏", "冶氏", "厉氏"
                )
                # -------------------------------------------------------------------

                # 检查用户对象是否有Group属性
                $hasGroupProperty = $false
                if ($users.Count -gt 0) {
                    $firstUser = $users[0]
                    $hasGroupProperty = $firstUser.PSObject.Properties.Name -contains "Group"
                }

                # 收集并创建所需的组
                if ($hasGroupProperty) {
                    $groups = $users | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Group) } | Select-Object -ExpandProperty Group -Unique
                    $result.TotalGroups = $groups.Count

                    if ($groups.Count -gt 0) {
                        foreach ($group in $groups) {
                            try {
                                # 检查组是否存在
                                $existingGroup = Get-ADGroup -Identity $group -ErrorAction Stop
                            }
                            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                                # 组不存在，尝试创建
                                try {
                                    if ([string]::IsNullOrWhiteSpace($NameOU)) {
                                        $result.ErrorLogs += "无法确定组 $group 的OU，跳过创建"
                                        continue
                                    }
                                    
                                    # 验证OU是否存在
                                    if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$NameOU'" -ErrorAction SilentlyContinue) -and
                                        -not (Get-ADObject -Filter "DistinguishedName -eq '$NameOU'" -ErrorAction SilentlyContinue)) {
                                        $result.ErrorLogs += "组的目标OU不存在: $NameOU，无法创建组 $group"
                                        continue
                                    }
                                    
                                    # 创建组参数
                                    $groupParams = @{
                                        Name            = $group
                                        SamAccountName  = $group
                                        GroupCategory   = "Security"
                                        GroupScope      = "Global"
                                        Path            = $NameOU
                                        Description     = "请更改组描述（自动创建）: $group"
                                        ErrorAction     = "Stop"
                                    }
                                    
                                    New-ADGroup @groupParams
                                    $result.CreatedGroups++
                                    $result.CreatedGroupsDetails += $group  # 记录创建成功的组名
                                }
                                catch {
                                    $result.ErrorLogs += "创建组 $group 失败: $($_.Exception.Message)"
                                }
                            }
                            catch {
                                $result.ErrorLogs += "检查组 $group 时发生错误: $_"
                            }
                        }
                    }
                }

                # 创建用户
                foreach ($user in $users) {
                    try {
                        # 检查用户是否已存在
                        $existingUser = Get-ADUser -Identity $user.SamAccountName -Properties DisplayName -ErrorAction Stop
                        
                        # 记录已存在的用户信息
                        $result.ExistingUsers += [PSCustomObject]@{
                            SamAccountName = $existingUser.SamAccountName
                            DisplayName = $existingUser.DisplayName
                        }
                        
                        $result.ErrorLogs += "用户 $($user.SamAccountName) 已存在，跳过创建"
                        $result.SkippedUsers++
                        continue
                    }
                    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                        # 用户不存在，继续创建流程
                    }
                    catch {
                        $result.ErrorLogs += "检查用户 $($user.SamAccountName) 时发生错误: $_"
                        $result.SkippedUsers++
                        continue
                    }
                    
                    # 字段验证
                    $requiredFields = 'SamAccountName'
                    if ($missingFields = $requiredFields | Where-Object { [string]::IsNullOrWhiteSpace($user.$_) }) {
                        $result.ErrorLogs += "用户记录缺少必填字段: $($missingFields -join ', ')，跳过"
                        $result.SkippedUsers++
                        continue
                    }

                    # 密码生成
                    $plainPassword = if (-not [string]::IsNullOrWhiteSpace($user.Password)) { 
                        $user.Password 
                    } else { 
                        "P@ssw0rd$(Get-Random -Minimum 1000 -Maximum 9999)" 
                    }
                    $securePassword = ConvertTo-SecureString $plainPassword -AsPlainText -Force

                    #账号过期日期处理
                    $accountExpirationDate = $null
                    if (-not [string]::IsNullOrWhiteSpace($user.AccountExpirationDate)) {
                        try {
                            $date = [DateTime]::MinValue
                            # 优先尝试默认格式解析
                            $parsed = [DateTime]::TryParse($user.AccountExpirationDate, [ref]$date)
                            
                            if (-not $parsed) {
                                # 尝试多种常见日期格式解析
                                $formats = @(
                                    "yyyy.MM.dd", "MM/dd/yyyy", "dd/MM/yyyy",
                                    "yyyy-MM-dd", "yyyyMMdd", "yyyy/MM/dd",
                                    "yyyy年MM月dd日"
                                )
                                foreach ($format in $formats) {
                                    if ([DateTime]::TryParseExact(
                                        $user.AccountExpirationDate, 
                                        $format, 
                                        [System.Globalization.CultureInfo]::InvariantCulture, 
                                        [System.Globalization.DateTimeStyles]::None, 
                                        [ref]$date
                                    )) {
                                        $parsed = $true
                                        break
                                    }
                                }
                            }
                            
                            if ($parsed) {
                                # AD需“过期日期+1天”（例：用户输入2024-10-05，AD设为2024-10-06，10-06午夜过期）
                                $accountExpirationDate = $date.Date.AddDays(1)
                            } else {
                                $result.ErrorLogs += "无法解析过期日期: $($user.AccountExpirationDate)，将不设置过期日期"
                            }
                        }
                        catch {
                            $result.ErrorLogs += "解析过期日期失败: $($user.AccountExpirationDate)，将不设置过期日期"
                        }
                    }

                    #复姓拆分逻辑
                    $fullName = $null
                    if (-not [string]::IsNullOrWhiteSpace($user.Surname) -and -not [string]::IsNullOrWhiteSpace($user.GivenName)) {
                        $surname = $user.Surname.Trim()
                        $givenName = $user.GivenName.Trim()
                    } else {
                        # 从DisplayName或Name字段获取完整姓名（CSV需至少有一个）
                        if (-not [string]::IsNullOrWhiteSpace($user.DisplayName)) {
                            $fullName = $user.DisplayName.Trim()
                        } elseif (-not [string]::IsNullOrWhiteSpace($user.Name)) {
                            $fullName = $user.Name.Trim()
                        } else {
                            $result.ErrorLogs += "用户 $($user.SamAccountName) 缺少姓名字段（DisplayName/Name），无法拆分姓和名"
                            $result.SkippedUsers++
                            continue
                        }

                        # 复姓识别与拆分
                        if ($fullName.Length -ge 2 -and $doubleSurnames -contains $fullName.Substring(0, 2)) {
                            $surname = $fullName.Substring(0, 2)
                            $givenName = if ($fullName.Length -gt 2) { $fullName.Substring(2).Trim() } else { "" }
                        } else {
                            $surname = $fullName.Substring(0, 1)
                            $givenName = if ($fullName.Length -gt 1) { $fullName.Substring(1).Trim() } else { $fullName }
                        }
                    }
                    # -------------------------------------------------------------------

                    # 获取当前域信息
                    $domain = Get-ADDomain
                    
                    # 用户参数
					$userParams = @{
						SamAccountName        = $user.SamAccountName
						UserPrincipalName     = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { "$($user.SamAccountName)@$($domain.DNSRoot)" }
						Name                  = $user.SamAccountName
						DisplayName           = if ($user.DisplayName) { $user.DisplayName } else { if ($fullName) { $fullName } else { $user.SamAccountName } }
						Surname               = $surname
						GivenName             = $givenName
						Path                  = $NameOU
						AccountPassword       = $securePassword
						Enabled               = $true
						ChangePasswordAtLogon = $true
						Description           = $user.Description
						EmailAddress          = $user.EmailAddress
						OfficePhone           = if (-not [string]::IsNullOrWhiteSpace($user.Telephone)) { $user.Telephone } else { $user.Phone }
						ErrorAction           = "Stop"
					}

                    # 只有解析成功时才添加过期日期参数
                    if ($accountExpirationDate) {
                        $userParams['AccountExpirationDate'] = $accountExpirationDate
                    }

                    # 清理空参数
                    $keysToRemove = $userParams.Keys | Where-Object { 
                        $null -eq $userParams[$_] -or [string]::IsNullOrWhiteSpace($userParams[$_]) 
                    }
                    $keysToRemove | ForEach-Object { $userParams.Remove($_) }

                    # 创建用户
                    try {
                        New-ADUser @userParams
                        $result.CreatedUsers++
                        
                        # 记录创建成功的用户详情
                        $result.CreatedUsersDetails += [PSCustomObject]@{
                            Index = $result.CreatedUsers  # 记录序号
                            SamAccountName = $user.SamAccountName
                            DisplayName = $userParams.DisplayName
                            Surname = $surname
                            GivenName = $givenName
							Email = $user.EmailAddress
							Telephone = if (-not [string]::IsNullOrWhiteSpace($user.Telephone)) { $user.Telephone } else { $user.Phone }
                            Password = $plainPassword
                        }
                        
                        # 将用户添加到组
                        if ($hasGroupProperty -and -not [string]::IsNullOrWhiteSpace($user.Group)) {
                            if (Get-ADGroup -Identity $user.Group -ErrorAction SilentlyContinue) {
                                Add-ADGroupMember -Identity $user.Group -Members $user.SamAccountName -ErrorAction Stop
                            } else {
                                $result.ErrorLogs += "组 $($user.Group) 不存在，无法将用户 $($user.SamAccountName) 添加到该组"
                            }
                        }
                    } catch {
                        $result.ErrorLogs += "创建用户 $($user.SamAccountName) 失败: $($_.Exception.Message)"
                        $result.SkippedUsers++
                    }
                }
                
                return $result
            } -ArgumentList $usersJson, $script:currentOU -ErrorAction Stop

            # 拼接结果信息
            $msg = @"
批量创建完成！
========================================
总用户数：$($result.TotalUsers)
成功创建：$($result.CreatedUsers)
跳过用户：$($result.SkippedUsers)
创建组数：$($result.CreatedGroups)/$($result.TotalGroups)
========================================

"@

            # 添加成功创建的组
            if ($result.CreatedGroups -gt 0) {
                $msg += "成功创建的组：`n$($result.CreatedGroupsDetails -join '、')`n`n"
				LoadGroupList
            }

            # 添加已存在的用户
            if ($result.ExistingUsers.Count -gt 0) {
                $msg += "已存在的用户：`n"
                $result.ExistingUsers | ForEach-Object {
                    $msg += "$($_.SamAccountName)（$($_.DisplayName)）`n"
                }
                $msg += "`n"
            }

            # 添加创建成功的用户
            if ($result.CreatedUsers -gt 0) {
                $msg += "创建成功的用户（初始密码）：`n"
                $result.CreatedUsersDetails | ForEach-Object {
                    $msg += "$($_.SamAccountName) | 姓名：$($_.DisplayName) | 密码：$($_.Password)`n"
                }
                $msg += "`n"
            }

            # 添加错误信息
			if ($result.ErrorLogs.Count -gt 0) {
				$msg += "`n错误信息：`n$($result.ErrorLogs -join "`r`n")`n"
			}

            # 显示结果对话框
            [System.Windows.Forms.MessageBox]::Show($msg, "批量创建结果", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

            # 刷新用户列表
            LoadUserList
            $script:connectionStatus = "批量创建用户操作已完成"
        }
        catch {
            $script:connectionStatus = "批量创建用户失败: $($_.Exception.Message)"
            $statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
            [System.Windows.Forms.MessageBox]::Show("执行批量创建时发生错误: $($_.Exception.Message)", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}


function ExportCSVUsers {
    # ---------------------- 1. 通用前置校验 + 初始化 ----------------------
    $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $script:statusOutputLabel.Text = "开始执行导出逻辑，初始化校验..."

    if (-not $script:domainContext) {
        $script:statusOutputLabel.Text = "【错误】请先连接到域控"
        [System.Windows.Forms.MessageBox]::Show("请先连接到域控", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    if (-not $script:remoteSession -or $script:remoteSession.State -ne "Opened") {
        $script:statusOutputLabel.Text = "【错误】远程会话已断开"
        [System.Windows.Forms.MessageBox]::Show("远程会话已断开，请重新连接域控", "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $script:statusOutputLabel.Text = "前置校验通过，开始判断场景..."

    # ---------------------- 2. 智能识别场景 ----------------------
    $sceneConfig = [PSCustomObject]@{
        SceneName    = ""          # Combined/Selected/Searched/All
        TargetSams   = @()         # 强制初始化为数组，避免null
        SceneDesc    = ""          # 场景描述
        ValidatePass = $true       # 校验结果
    }

    $searchText = $script:textSearch.Text.Trim()
    $filteredCount = $script:filteredUsers.Count
    $allCount = $script:allUsers.Count
    $hasSearchFilter = (-not [string]::IsNullOrEmpty($searchText)) -or ($filteredCount -ne $allCount)
    $hasSelectedRows = $script:userDataGridView.SelectedRows.Count -gt 0

    $script:statusOutputLabel.Text = "筛选状态：搜索框='$searchText'，筛选后数=$filteredCount，全量数=$allCount"

    # 场景1：筛选结果中的多选（组合场景）
    if ($hasSearchFilter -and $hasSelectedRows) {
        $sceneConfig.SceneName = "Combined"
        $selectedSams = @()  # 强制数组
        foreach ($row in $script:userDataGridView.SelectedRows) {
            $samCell = $row.Cells | Where-Object { $_.OwningColumn.Name -eq "SamAccountName" }
            if ($samCell -and -not [string]::IsNullOrEmpty($samCell.Value)) {
                $selectedSams += $samCell.Value.Trim()
            }
        }
        $sceneConfig.TargetSams = $selectedSams
        $sceneConfig.SceneDesc = "筛选结果中的 $($sceneConfig.TargetSams.Count) 个选中用户"

        if ($sceneConfig.TargetSams.Count -eq 0) {
            $script:statusOutputLabel.Text = "【错误】选中行无有效用户数据"
            [System.Windows.Forms.MessageBox]::Show("选中行无有效用户数据，请重新选择", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $sceneConfig.ValidatePass = $false
        }
    }
    # 场景2：单纯多选（无筛选）
    elseif (-not $hasSearchFilter -and $hasSelectedRows) {
        $sceneConfig.SceneName = "Selected"
        $selectedSams = @()  # 强制数组
        foreach ($row in $script:userDataGridView.SelectedRows) {
            $samCell = $row.Cells | Where-Object { $_.OwningColumn.Name -eq "SamAccountName" }
            if ($samCell -and -not [string]::IsNullOrEmpty($samCell.Value)) {
                $selectedSams += $samCell.Value.Trim()
            }
        }
        $sceneConfig.TargetSams = $selectedSams
        $sceneConfig.SceneDesc = "全量数据中的 $($sceneConfig.TargetSams.Count) 个选中用户"
        
        if ($sceneConfig.TargetSams.Count -eq 0) {
            $script:statusOutputLabel.Text = "【错误】选中行无有效用户数据"
            [System.Windows.Forms.MessageBox]::Show("选中行无有效用户数据，请重新选择", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $sceneConfig.ValidatePass = $false
        }
    }
    # 场景3：检索量（有筛选，无选中行）
    elseif ($hasSearchFilter -and -not $hasSelectedRows) {
        $sceneConfig.SceneName = "Searched"
        $sceneConfig.TargetSams = @($script:filteredUsers | ForEach-Object {
            if (-not [string]::IsNullOrEmpty($_.SamAccountName)) {
                return $_.SamAccountName.Trim()
            }
        })
        $sceneConfig.SceneDesc = "检索结果（共 $($sceneConfig.TargetSams.Count) 个用户）"
        
        if ($sceneConfig.TargetSams.Count -eq 0) {
            $script:statusOutputLabel.Text = "【错误】检索结果为空"
            [System.Windows.Forms.MessageBox]::Show("当前检索结果为空，无法导出", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $sceneConfig.ValidatePass = $false
        }
    }
    # 场景4：全量
    else {
        $sceneConfig.SceneName = "All"
        $sceneConfig.TargetSams = @()
        $sceneConfig.SceneDesc = "当前OU下所有用户（共 $allCount 个）"
        
        if ($allCount -eq 0) {
            $script:statusOutputLabel.Text = "【错误】当前OU无用户数据"
            [System.Windows.Forms.MessageBox]::Show("当前OU下无用户数据，无法导出", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $sceneConfig.ValidatePass = $false
        }
    }

    if (-not $sceneConfig.ValidatePass) {
        $script:statusOutputLabel.Text = "【错误】场景校验失败"
        return
    }
    $script:statusOutputLabel.Text = "场景判断完成：$($sceneConfig.SceneDesc)"

    # ---------------------- 新增：选中1个用户时的询问弹窗（含Ctrl多选提示） ----------------------
    if (($sceneConfig.SceneName -eq "Combined" -or $sceneConfig.SceneName -eq "Selected") -and $sceneConfig.TargetSams.Count -eq 1) {
        $confirmMsg = @"
当前选中1个用户，是否继续导出？

提示：若需同时导出多个用户，可按住Ctrl键点击多个用户行进行多选。
或使用检索关键字，导出当前检索列表下的所有用户数据。
    `n若没有作任何选中，默认导出当前OU下所有用户数据。
"@
        $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "确认导出", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::No) {
            $script:statusOutputLabel.Text = "用户取消导出（选中1个用户）"
            return
        }
    }

    # ---------------------- 3. 文件保存对话框 ----------------------
    $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $fileDialog.Filter = "CSV文件 (*.csv)|*.csv|所有文件 (*.*)|*.*"
    $fileDialog.Title = "导出CSV（当前：$($sceneConfig.SceneDesc)）"
    $fileDialog.DefaultExt = "csv"
    $fileDialog.AddExtension = $true
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    $script:statusOutputLabel.Text = "等待用户选择保存路径..."
    if ($fileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        $script:statusOutputLabel.Text = "用户取消导出"
        return
    }
    $csvPath = $fileDialog.FileName
    $script:statusOutputLabel.Text = "选中路径：$($csvPath | Split-Path -Leaf)"

    # ---------------------- 4. 远程读取+本地导出） ----------------------
    try {
        $script:statusOutputLabel.Text = "开始远程读取AD数据..."

        # 确保参数正确传递
        $remoteUserData = Invoke-Command -Session $script:remoteSession -ScriptBlock {
            param($NameOU, $allUsersOU, $targetSams, $sceneName)
            
            # 强制转为数组，确保单元素场景正常
            $targetSams = @($targetSams | Where-Object { -not [string]::IsNullOrEmpty($_) })
            $remoteResult = [PSCustomObject]@{ UserData = $null; ErrorLogs = @() }

            # 检查AD模块
            if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
                $remoteResult.ErrorLogs += "域控未安装Active Directory模块"
                return $remoteResult
            }
            try { 
                Import-Module ActiveDirectory -ErrorAction Stop 
            }
            catch { 
                $remoteResult.ErrorLogs += "导入AD模块失败：$($_.Exception.Message)"
                return $remoteResult 
            }

            # 构建查询Filter
            $adProperties = @("SamAccountName", "DisplayName", "UserPrincipalName", "EmailAddress", "TelephoneNumber", "Description", "GivenName", "Surname", "AccountExpirationDate", "Enabled", "DistinguishedName", "Department", "Title", "OfficePhone", "LastLogonDate")
            $filter = "*"
            if ($targetSams.Count -gt 0) {
                $samConditions = $targetSams | ForEach-Object { "SamAccountName -eq '$_'" }
                $filter = $samConditions -join ' -or '
            }

            # 执行AD查询
            try {
                $getADParams = @{
                    Filter     = $filter
                    Properties = $adProperties
                    ErrorAction = "Stop"
                }
                if (-not [string]::IsNullOrEmpty($NameOU) -and -not $allUsersOU) {
                    $getADParams["SearchBase"] = $NameOU
                }

                $users = Get-ADUser @getADParams
                # 强制转换为数组，确保单个用户时也是数组
                $remoteResult.UserData = @($users | Select-Object $adProperties)
            }
            catch { 
                $remoteResult.ErrorLogs += "读取AD用户失败：$($_.Exception.Message)"
                return $remoteResult 
            }

            return $remoteResult
        } -ArgumentList $script:currentOU, $script:allUsersOU, $sceneConfig.TargetSams, $sceneConfig.SceneName -ErrorAction Stop

        # 远程结果处理
        $script:statusOutputLabel.Text = "远程数据读取完成，开始校验..."
        
        # 正确处理返回的用户数据计数
        $remoteUserCount = if ($remoteUserData.UserData) { 
            if ($remoteUserData.UserData -is [array]) {
                $remoteUserData.UserData.Count
            } else {
                1  # 单个对象
            }
        } else { 
            0 
        }

        # 处理远程错误
        if ($remoteUserData.ErrorLogs.Count -gt 0) {
            throw "远程读取失败：`n$($remoteUserData.ErrorLogs -join "`n")"
        }
        if (-not $remoteUserData.UserData -or $remoteUserCount -eq 0) {
            throw "未读取到 $($sceneConfig.SceneDesc) 的数据（远程返回空）"
        }

        # ---------------------- 5. 本地处理+导出（单用户处理） ----------------------
        $script:statusOutputLabel.Text = "开始本地格式化数据..."

        # 确保单个用户对象也能被正确处理
        $userDataArray = @($remoteUserData.UserData)
        $exportData = $userDataArray | ForEach-Object {
            [PSCustomObject]@{
                "用户名(SamAccountName)"    = $_.SamAccountName
                "显示名称"                  = $_.DisplayName
                "姓(Surname)"               = $_.Surname
                "名(GivenName)"             = $_.GivenName
                "邮箱地址"                  = $_.EmailAddress
                "电话(TelephoneNumber)"     = if ($_.TelephoneNumber) { $_.TelephoneNumber } else { "无" }
                "描述"                      = $_.Description
                "账号过期日期"              = if ($_.AccountExpirationDate) { $_.AccountExpirationDate.AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss") } else { "无" }
                "最后登录时间"              = if ($_.LastLogonDate) { $_.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "从未登录" }
                "账号状态"                  = if ($_.Enabled) { "启用" } else { "禁用" }
                "所属OU"                    = $_.DistinguishedName
                "部门"                      = $_.Department
                "职位"                      = $_.Title
                "办公电话"                  = $_.OfficePhone
            }
        }

        # 确保exportData是数组
        $exportData = @($exportData)
        $exportCount = $exportData.Count

        # 导出CSV
        $exportDir = [System.IO.Path]::GetDirectoryName($csvPath)
        if (-not (Test-Path -Path $exportDir -PathType Container)) {
            throw "导出目录不存在：$exportDir"
        }
        
        # 权限测试
        $testFile = [System.IO.Path]::Combine($exportDir, "tmp_export_test.tmp")
        New-Item -Path $testFile -ItemType File -Force | Out-Null
        Remove-Item -Path $testFile -Force | Out-Null

        # 正式导出
        $exportData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
        $script:statusOutputLabel.Text = "CSV导出完成！"

        # ---------------------- 6. 成功提示 ----------------------
        $successMsg = @"
导出成功！
========================================
导出内容：$($sceneConfig.SceneDesc)
文件路径：$csvPath
导出数量：$exportCount 个用户
========================================
是否立即打开文件查看？
"@
        if ([System.Windows.Forms.MessageBox]::Show($successMsg, "导出完成", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information) -eq "Yes") {
            Start-Process -FilePath $csvPath
        }
        $script:statusOutputLabel.Text = "导出成功：$($sceneConfig.SceneDesc)"
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    catch {
        $errorMsg = "导出 $($sceneConfig.SceneDesc) 失败：$($_.Exception.Message)"
        $script:statusOutputLabel.Text = "【错误】$errorMsg"
        $script:statusOutputLabel.ForeColor = [System.Drawing.Color]::DarkRed
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "导出错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}
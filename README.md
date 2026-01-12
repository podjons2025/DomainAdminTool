调整 Forms/Controls/ConnectionPanel.ps1 适配

$script:comboDomain.Items.AddRange(@(	
    [PSCustomObject]@{Name = "域控（广州）- serverAD.abc.com"; Server = "serverAD.abc.com"; SystemAccount= "abc\admin"; Password = "Abc123456"},
    [PSCustomObject]@{Name = "域控（上海）- abc03.abc01.com"; Server = "abc03.abc01.com"; SystemAccount= "abc01\administrator"; Password = "Password123"},
    [PSCustomObject]@{Name = "测试域控（北京）- serverAD3.abc03.com"; Server = "serverAD3.abc03.com"; SystemAccount= "abc03\admin"; Password = ""}		
))

Name = ""
Server = ""
SystemAccount= ""
Password = ""

<img width="1185" height="892" alt="image" src="https://github.com/user-attachments/assets/684a08d9-4c51-4008-9a3d-0898653ae4c6" />




Public Class Login_Main

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        If System.IO.File.Exists(Application.StartupPath & "\ERP Update.exe_1") Then
            System.IO.File.Delete(Application.StartupPath & "\ERP Update.exe")
            System.IO.File.Copy(Application.StartupPath & "\ERP Update.exe_1", Application.StartupPath & "\ERP Update.exe", False)
            System.IO.File.Delete(Application.StartupPath & "\ERP Update.exe_1")
        End If

        Dim run_file As String = Application.StartupPath

        ''DB 접속
        'DBConnect()

        'Dim strSQL As String = "select user_pass, user_authority, user_vip, user_name from user_manager where user_id = '" & USER_ID.Text & "'"

        'Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        'Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader
        'Dim find_ok As Boolean = False

        'Do While sqlDR.Read
        '    find_ok = True
        '    'PASSWORD가 검사
        '    If sqlDR("USER_PASS") = USER_PASS.Text Then
        '        If cb_RunProgram.Text = "Repair System" Then
        '            registryEdit.WriteRegKey("Software\Yujin\Repair System\Login", "User Name", sqlDR("USER_NAME"))
        '            registryEdit.WriteRegKey("Software\Yujin\Repair System\Login", "User_Authority", sqlDR("USER_AUTHORITY"))
        '        End If
        '        Exit Do
        '    Else
        '        MsgBox("Password가 다릅니다.", MsgBoxStyle.Information, "YUJIN Manufacturing Management System")
        '        USER_PASS.SelectAll()
        '        USER_PASS.Focus()
        '        sqlDR.Close()
        '        DBClose()
        '        Exit Sub
        '    End If
        'Loop

        'sqlDR.Close()

        'DBClose()

        'If find_ok = False Then
        '    MsgBox("ID를 찾을 수 없습니다.", MsgBoxStyle.Information, "YUJIN Manufacturing Management System")
        '    USER_ID.SelectAll()
        '    USER_ID.Focus()
        '    Exit Sub
        'End If

        If cb_RunProgram.Text = "Repair System" Then
            '마지막 Login된 ID를 기록
            registryEdit.WriteRegKey("Software\Yujin\Repair System\Login", "User ID", USER_ID.Text)
            registryEdit.WriteRegKey("Software\Yujin\Repair System\Login", "Site", cb_RunProgram.Text)

            If CB_PASS_SAVE.Checked Then
                registryEdit.WriteRegKey("Software\Yujin\Repair System\Login", "Pass Save", "Yes")
                registryEdit.WriteRegKey("Software\Yujin\Repair System\Login", "Pass", USER_PASS.Text)
            Else
                registryEdit.WriteRegKey("Software\Yujin\Repair System\Login", "Pass Save", "No")
                registryEdit.WriteRegKey("Software\Yujin\Repair System\Login", "Pass", String.Empty)
            End If

            run_file += "\Repair_System.exe"
        End If

        'Shell(run_file, AppWinStyle.NormalFocus)
        CreateObject("WScript.Shell").Run(Chr(34) & run_file & Chr(34), True)

        Me.Close()

    End Sub

    Private Sub MMSLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '마지막 Login된 ID를 불러와서 USER_ID에 기록
        cb_RunProgram.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Login", "Site", "Repair System")

        If cb_RunProgram.Text = "Repair System" Then
            USER_ID.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Login", "User ID", "ADMIN")

            If Not USER_ID.Text = "" Then USER_PASS.Focus()

            If registryEdit.ReadRegKey("Software\Yujin\Repair System\Login", "Pass Save", "No") = "Yes" Then
                CB_PASS_SAVE.Checked = True
                USER_PASS.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Login", "Pass", String.Empty)
            End If
        End If

    End Sub

    Private Sub CONN_SITE_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_RunProgram.SelectedIndexChanged

        USER_ID.Enabled = True
        USER_PASS.Enabled = True

        If cb_RunProgram.Text = "Repair System" Then

            USER_ID.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Login", "User ID", "ADMIN")

            If Not USER_ID.Text = "" Then USER_PASS.Focus()

            If registryEdit.ReadRegKey("Software\Yujin\Repair System\Login", "Pass Save", "No") = "Yes" Then
                CB_PASS_SAVE.Checked = True
                USER_PASS.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Login", "Pass", String.Empty)
            End If
        End If

    End Sub
End Class

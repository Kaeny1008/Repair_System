Public Class SQL_Connection

    Private Sub SQL_Connection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        TB_IP.Text = registryEdit.ReadRegKey("Software\Yujin\MMS", "Server.IP", "192.168.0.254")
        TB_PORT.Text = registryEdit.ReadRegKey("Software\Yujin\MMS", "Server.PORT", "3306")
        TB_ID.Text = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS", "Server.ID", "yujin")
        TB_PW.Text = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS", "Server.PSWD", "Park1052!")
        UPDATE_CHECK.Text = CInt(registryEdit.ReadRegKey("Software\Yujin\MMS", "Update Check", "10000")) / 1000

    End Sub

    Private Sub BTN_SAVE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_SAVE.Click

        registryEdit.WriteRegKey("Software\Yujin\MMS", "Server.IP", TB_IP.Text)
        registryEdit.WriteRegKey("Software\Yujin\MMS", "Server.PORT", TB_PORT.Text)
        registryEdit.WriteRegKey("Software\Yujin\MMS_MMPS", "Server.ID", TB_ID.Text)
        registryEdit.WriteRegKey("Software\Yujin\MMS_MMPS", "Server.PSWD", TB_PW.Text)
        registryEdit.WriteRegKey("Software\Yujin\MMS", "Update Check", CInt(UPDATE_CHECK.Text) * 1000)
        serverID = TB_ID.Text
        serverPORT = TB_PORT.Text
        serverIP = TB_IP.Text
        serverPSWD = TB_PW.Text

        DBConnect()
        If DBConnect1.State = 1 Then
            DBClose()
            MsgBox("접속성공", MsgBoxStyle.Information, "자산관리")
            Me.Dispose()
        End If

    End Sub
End Class
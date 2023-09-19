Imports MySql.Data.MySqlClient

Module MySQL

    Public registryEdit As New RegistryEdit.RegReadWrite '레지스트리 편집
    Public DBConnect1 As MySqlConnection
    Public serverIP As String = registryEdit.ReadRegKey("Software\Yujin\Repair System", "Server.IP", "192.168.0.222")
    Public serverPORT As String = registryEdit.ReadRegKey("Software\Yujin\Repair System", "Server.PORT", "10522")
    Public serverID As String = registryEdit.ReadRegKey("Software\Yujin\Repair System", "Server.ID", "yujin_REPAIR")
    Public serverPSWD As String = registryEdit.ReadRegKey("Software\Yujin\Repair System", "Server.PSWD", "Dbwlswjswk1!")

    'DB 연결 함수
    Public Sub DBConnect()

        DBConnect1 = New MySqlConnection
        DBConnect1.ConnectionString = "Database=repair_system" &
                                       ";Data Source=" & serverIP &
                                       ";PORT=" & serverPORT &
                                       ";User Id=" & serverID &
                                       ";Password=" & serverPSWD &
                                       ";allow user variables=true"

        Try
            DBConnect1.Open()
            'DBConnect1 연결되어있지 않다면
            'If Not DBConnect1.State = ConnectionState.Open Then
            '        MessageBox.Show("DB 연결 실패", "DB 테스트", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'End If
            Dim strSql As String = "SET Names euckr;"
            Dim sqlCmd As New MySqlCommand(strSql, DBConnect1)
        Catch ex As Exception
            MessageBox.Show("DB 연결 실패", "DB 테스트", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    'DB 종료 함수
    Public Sub DBClose()
        If DBConnect1.State = ConnectionState.Open Then
            DBConnect1.Close()
        End If
    End Sub
End Module

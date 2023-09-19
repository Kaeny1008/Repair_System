Imports MySql.Data.MySqlClient

Public Class frm_Main

    Dim pg_ver As String = String.Empty
    Dim form_msgbox_string As String = "Update Checker"
    Dim update_check_time As Integer = CInt(registryEdit.ReadRegKey("Software\Yujin\Repair System", "Update Check", "10000"))

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        VER_OPEN()

        Timer1.Interval = 1000
        Timer1.Enabled = True
        Timer2.Interval = update_check_time
        Timer2.Enabled = True
        Timer3.Interval = 60000
        Timer3.Enabled = True

        Me.Hide()
        Me.Visible = False
        Me.Opacity = 0

    End Sub

    Private Sub VER_OPEN()

        pg_ver = String.Empty

        Dim ReadOk As Integer
        Dim openFileDialog1 As New OpenFileDialog
        '새로운파일버젼 확인
        openFileDialog1.FileName = Application.StartupPath & "\CheckVer.ini"

        Dim fs As System.IO.FileStream
        Dim sr As System.IO.StreamReader
        fs = System.IO.File.Open(openFileDialog1.FileName, IO.FileMode.Open) ' 파일 열기
        sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결

        Dim str As String
        Dim NewVer() As String

        'While sr.EndOfStream <> False ' 끝까지 읽기
        While sr.Peek() >= 0
            str = sr.ReadLine() ' 한줄씩 읽기
            If str = "[AUTO UPDATE]" Then ReadOk = 1 '파일읽기 시작
            If str = "[/]" Then ReadOk = 0 '파일읽기 종료

            NewVer = Split(str, "=")

            If ReadOk = 1 And NewVer.Length > 1 Then '읽시 시작이 Ok면
                If pg_ver = String.Empty Then
                    '파일명,폴더명,버전
                    pg_ver = NewVer(0) & "::" & NewVer(2) & "::" & NewVer(1)
                Else
                    '파일명,폴더명,버전
                    pg_ver += "//" & NewVer(0) & "::" & NewVer(2) & "::" & NewVer(1)
                End If
                'Console.WriteLine(pg_ver)
            End If
        End While
        sr.Close()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If CDate(Format(Now, "yyyy-MM-dd HH:mm:ss")) > CDate(after_alarm) Then
            If next_alarm = True Then Timer2.Enabled = True
        End If

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        VER_OPEN()
        Console.WriteLine("PC에 저장된 버전 파일을 열었습니다.")

        '업데이트 체크 시간을 강제로 변경... 차후에 다시 돌려놓자
        '10초마다 체크는 비효율적일듯.
        'registryEdit.WriteRegKey("Software\Yujin\MMS", "Update Check", 60000)

        'PG_STATUS.Text = "Program Ver. 비교가 시작 되었습니다."

        DBConnect()

        Dim strSql As String = "select file_name, folder_name, upload_date from setting_program_manager order by folder_name, file_code"

        Dim sqlCmd As New MySqlCommand(strSql, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader
        Dim module_name As String = String.Empty

        Console.WriteLine("버전 비교가 시작되었습니다.")
        Do While sqlDR.Read
            Dim pg_ver2() As String = Split(pg_ver, "//")
            Dim name_find As Boolean = False
            For i = 0 To UBound(pg_ver2)
                Dim pg_ver3() As String = Split(pg_ver2(i), "::") '::로 분할하여 업데이트 시간이 동일한지 검사
                If pg_ver3(0) = sqlDR("file_name") Then '파일명이 동일하다면
                    'Console.WriteLine(pg_ver3(2) = Format(sqlDR("upload_date"), "yyyy-MM-dd HH:mm:ss"))
                    If Not pg_ver3(2) = Format(sqlDR("upload_date"), "yyyy-MM-dd HH:mm:ss") Then
                        '업데이트 시간이 동일하지 않다면...
                        If module_name = String.Empty Then
                            module_name = pg_ver3(0)
                        Else
                            module_name += vbCrLf & "               " & pg_ver3(0)
                        End If
                    End If
                    name_find = True
                    Exit For
                End If
            Next
            If name_find = False Then
                If module_name = String.Empty Then
                    module_name = sqlDR("file_name")
                Else
                    module_name += vbCrLf & "               " & sqlDR("FILE_NAME")
                End If
            End If
        Loop
        DBClose()
        Console.WriteLine("버전 비교가 종료 되었습니다.")

        'PG_STATUS.Text = "Program Ver. 비교가 완료 되었습니다."

        If Not module_name = String.Empty Then
            Console.WriteLine("업데이트된 내용 : " & module_name)
            UPDATE_ALARM(module_name)
        End If

    End Sub

    Dim after_alarm As String = Format(Now, "yyyy-MM-dd HH:mm:ss")
    Dim next_alarm As Boolean = False

    Private Sub UPDATE_ALARM(ByVal module_name As String)

        Timer2.Enabled = False
        next_alarm = False

        Dim message As String = "프로그램 업데이트가 있습니다." & vbCrLf &
                                "업데이트 하지 않고 사용시 심각한 오류가 발생할 수 있습니다." & vbCrLf &
                                "지금 업데이트 하시겠습니까?" & vbCrLf &
                                "(해당 모든 프로그램이 종료됩니다.)" & vbCrLf &
                                "File Name : " & module_name

        Using form = New Form() With {.TopMost = True} ' 최상위로 포커스
            If MessageBox.Show(form, message, form_msgbox_string, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                after_alarm = Format(DateAdd(DateInterval.Hour, 1, Now), "yyyy-MM-dd HH:mm:ss")
                next_alarm = True '재알림을 방지하기 위해!!
                MsgBox("업데이트가 1시간 연장 되었습니다." & vbCrLf &
                   "1시간 뒤 알람이 재발생합니다.", MsgBoxStyle.Information, form_msgbox_string)
            Else
                MsgBox("프로그램이 종료되고 업데이트가 실행 됩니다.", MsgBoxStyle.Information, form_msgbox_string)
                Try
                    Dim myProgram As String = "ERP Update, YJ Login, Repair_System"
                    Dim myProgram2() As String = Split(myProgram, ", ")
                    For i = 0 To UBound(myProgram2)
                        Dim pcProcess() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName(myProgram2(i))
                        If pcProcess.Length > 0 Then
                            pcProcess(0).Kill()
                        End If
                    Next
                    'Shell(Application.StartupPath & "\ERP Update.exe")
                    'Shell(Application.StartupPath & "\ERP Update.exe", AppWinStyle.NormalFocus)
                    CreateObject("WScript.Shell").Run(Chr(34) & Application.StartupPath & "\ERP Update.exe" & Chr(34), True)
                    Me.Close()
                Catch ex As Exception
                    MsgBox(ex.Message & vbCrLf & Application.StartupPath & "\ERP Update.exe", MsgBoxStyle.Critical, form_msgbox_string)
                    Me.Dispose()
                End Try
            End If
        End Using

        'If MsgBox("프로그램 업데이트가 있습니다." & vbCrLf &
        '          "업데이트 하지 않고 사용시 심각한 오류가 발생할 수 있습니다." & vbCrLf &
        '          "지금 업데이트 하시겠습니까?" & vbCrLf &
        '          "(해당 모든 프로그램이 종료됩니다.)" & vbCrLf &
        '          "File Name : " & module_name, MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.No Then
        '    after_alarm = Format(DateAdd(DateInterval.Hour, 1, Now), "yyyy-MM-dd HH:mm:ss")
        '    next_alarm = True '재알림을 방지하기 위해!!
        '    MsgBox("업데이트가 1시간 연장 되었습니다." & vbCrLf &
        '           "1시간 뒤 알람이 재발생합니다.", MsgBoxStyle.Information, form_msgbox_string)
        'Else
        '    MsgBox("프로그램이 종료되고 업데이트가 실행 됩니다.", MsgBoxStyle.Information, form_msgbox_string)
        '    Try
        '        Dim myProgram As String = "YJ MMS, YJ MMS Asset, YJ MMS Label Printer, YJ MMS 3rd, YJ MMS MMPS, YJ MMS MMS"
        '        Dim myProgram2() As String = Split(myProgram, ", ")
        '        For i = 0 To UBound(myProgram2)
        '            Dim pcProcess() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName(myProgram2(i))
        '            If pcProcess.Length > 0 Then
        '                pcProcess(0).Kill()
        '            End If
        '        Next

        '        Shell(Application.StartupPath & "\ERP Update.exe")
        '        Me.Close()
        '    Catch ex As Exception
        '        MsgBox(ex.Message, MsgBoxStyle.Critical, form_msgbox_string)
        '    End Try
        'End If

    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick

        Timer2.Interval = CInt(registryEdit.ReadRegKey("Software\Yujin\Repair System", "Update Check", "10000"))

    End Sub
End Class

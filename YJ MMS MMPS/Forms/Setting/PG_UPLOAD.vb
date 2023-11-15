Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class PG_UPLOAD

    Private Sub PG_UPLOAD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Me.WindowState = FormWindowState.Maximized
        GRID_SETTING()

        '최신 버전체크 파일을 다운로드 및 그리드에 표시
        FILE_LIST_LOAD()

    End Sub

    Private Sub GRID_SETTING()

        With FILE_LIST
            .AllowEditing = True
            .AllowFiltering = True
            .AllowSorting = AllowSortingEnum.SingleColumn
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 7
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            FILE_LIST(0, 0) = "No"
            FILE_LIST(0, 1) = "Division"
            FILE_LIST(0, 2) = "File Name"
            FILE_LIST(0, 3) = "Folder Name"
            FILE_LIST(0, 4) = "Upload Date/Ttime"
            FILE_LIST(0, 5) = "Upload ID"
            FILE_LIST(0, 6) = "비고"
            .Cols(2).ComboList = "..."
            '.Cols(4).DataType = GetType(DateAndTime)
            '.Cols(4).Format = "9999-99-99 99:99:99"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

    End Sub

    Private Sub FILE_LIST_LOAD()

        FILE_LIST.Rows.Count = 1

        DBConnect_HO()

        Dim strSql As String = "select FILE_CODE, FILE_NAME, FOLDER_NAME, UPLOAD_DATE, WRITE_ID, ifnull(UPLOAD_NOTE, '') as UPLOAD_NOTE from PROGRAM_UPDATE order by FOLDER_NAME, FILE_CODE"

        Dim sqlCmd As New MySqlCommand(strSql, DBConnect2)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_string As String = String.Empty
            insert_string = FILE_LIST.Rows.Count & vbTab & _
                            sqlDR("FILE_CODE") & vbTab & _
                            sqlDR("FILE_NAME") & vbTab & _
                            sqlDR("FOLDER_NAME") & vbTab & _
                            Format(sqlDR("UPLOAD_DATE"), "yyyy-MM-dd HH:mm:ss") & vbTab & _
                            sqlDR("WRITE_ID") & vbTab & _
                            sqlDR("UPLOAD_NOTE")
            FILE_LIST.AddItem(insert_string)
        Loop
        sqlDR.Close()

        DBClose_HO()

        FILE_LIST.AutoSizeCols()

    End Sub

    Private Function FILE_LIST_BEFORE_LOAD(ByVal col_name As String, ByVal file_code As String) As String

        FILE_LIST_BEFORE_LOAD = String.Empty

        DBConnect_HO()

        Dim strSql As String = "select ifnull(" & col_name & ", '') as " & col_name & " from PROGRAM_UPDATE"
        strSql += " where FILE_CODE = '" & file_code & "'"

        Dim sqlCmd As New MySqlCommand(strSql, DBConnect2)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            FILE_LIST_BEFORE_LOAD = sqlDR(col_name)
        Loop
        sqlDR.Close()

        DBClose_HO()

        Return FILE_LIST_BEFORE_LOAD

    End Function

    Private Sub FILE_LIST_AfterEdit(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles FILE_LIST.AfterEdit

        If IsNumeric(FILE_LIST(e.Row, 0)) = False Then Exit Sub

        Dim find_col_name As String = String.Empty
        Select Case e.Col
            Case 1
                find_col_name = "FILE_CODE"
            Case 2
                find_col_name = "FILE_NAME"
            Case 3
                find_col_name = "FOLDER_NAME"
            Case 4
                find_col_name = "UPLOAD_DATE"
            Case 5
                find_col_name = "WRITE_ID"
            Case 6
                find_col_name = "UPLOAD_NOTE"
        End Select

        If FILE_LIST_BEFORE_LOAD(find_col_name, FILE_LIST(e.Row, 1)) = FILE_LIST(e.Row, e.Col) = False Then
            If Not FILE_LIST(e.Row, 0).ToString = "N" Then
                FILE_LIST.Rows(e.Row).StyleNew.ForeColor = Color.Red
                FILE_LIST(e.Row, 0) = "M"
                FILE_LIST(e.Row, 4) = Format(Now, "yyyy-MM-dd HH:mm:ss")
                FILE_LIST(e.Row, 5) = MainForm.login_user
            End If
        Else
            FILE_LIST.Rows(e.Row).StyleNew.ForeColor = Color.Black
            FILE_LIST(e.Row, 0) = "0"
            For i = 1 To FILE_LIST.Rows.Count - 1
                If IsNumeric(FILE_LIST(i, 0)) = True Then
                    FILE_LIST(i, 0) = i
                End If
            Next
        End If

    End Sub

    Private Sub FILE_LIST_CellButtonClick(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles FILE_LIST.CellButtonClick

        Dim openFile As New System.Windows.Forms.OpenFileDialog

        With openFile
            .Filter = "ALL Files (*.*)|*.*"
            .AddExtension() = True
        End With

        openFile.ShowDialog()

        If openFile.FileName = "" Then Exit Sub

        FILE_LIST.Rows(e.Row).StyleNew.ForeColor = Color.Blue

        If IsNumeric(FILE_LIST(e.Row, 0)) Then FILE_LIST(e.Row, 0) = "M"

        Dim TEMP_FILE_Folder As String = Application.StartupPath & "\TEMP_FILE"

        '좌표폴더가 없으면 만들기
        If My.Computer.FileSystem.DirectoryExists(TEMP_FILE_Folder) = False Then
            My.Computer.FileSystem.CreateDirectory(TEMP_FILE_Folder)
        End If

        '선택한 파일을 임시폴더에 저장
        My.Computer.FileSystem.CopyFile(openFile.FileName, TEMP_FILE_Folder & "\" & openFile.FileName.Split("\")(UBound(openFile.FileName.Split("\"))), True)

        FILE_LIST(e.Row, 2) = openFile.FileName.Split("\")(UBound(openFile.FileName.Split("\")))

        If Not FILE_LIST(e.Row, 0) Like "*,U" Then
            FILE_LIST(e.Row, 0) += ",U"
        End If

        FILE_LIST(e.Row, 4) = Format(Now, "yyyy-MM-dd HH:mm:ss")
        FILE_LIST(e.Row, 5) = MainForm.login_user

        FILE_LIST.AutoSizeCols()

    End Sub

    Private Sub FILE_LIST_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles FILE_LIST.MouseClick

        If e.Button = Windows.Forms.MouseButtons.Right And FILE_LIST.MouseRow < 1 Then
            FILE_LIST.Row = FILE_LIST.MouseRow
            BTN_ADD.Enabled = False
            BTN_DELETE.Enabled = False
            GRID_MENU.Show(FILE_LIST, New Point(e.X, e.Y))
        ElseIf e.Button = Windows.Forms.MouseButtons.Right And FILE_LIST.MouseRow > 0 Then
            FILE_LIST.Row = FILE_LIST.MouseRow
            BTN_ADD.Enabled = True
            BTN_DELETE.Enabled = True
            GRID_MENU.Show(FILE_LIST, New Point(e.X, e.Y))
        End If

    End Sub

    Private Sub FILE_LIST_RowColChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles FILE_LIST.RowColChange

        If FILE_LIST.Rows.Count = 1 Then Exit Sub
        If FILE_LIST.Row < 1 Then Exit Sub

        Select Case FILE_LIST.Col
            Case 1
                If IsNumeric(FILE_LIST(FILE_LIST.Row, 0)) Then
                    FILE_LIST.AllowEditing = False
                Else
                    FILE_LIST.AllowEditing = True
                End If
            Case Else
                FILE_LIST.AllowEditing = True
        End Select

    End Sub

    Private Sub BTN_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_ADD.Click

        FILE_LIST.AddItem("N" & vbTab & String.Empty & vbTab & String.Empty & vbTab & String.Empty & vbTab & Format(Now, "yyyy-MM-dd HH:mm:ss") & vbTab & MainForm.login_user, FILE_LIST.Row + 1)
        FILE_LIST.TopRow = FILE_LIST.Rows.Count - 1
        FILE_LIST.Rows(FILE_LIST.Row + 1).StyleNew.ForeColor = Color.Blue

        For i = 1 To FILE_LIST.Rows.Count - 1
            If IsNumeric(FILE_LIST(i, 0)) = True Then
                FILE_LIST(i, 0) = i
            End If
        Next

    End Sub

    Private Sub BTN_DELETE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_DELETE.Click

        FILE_LIST(FILE_LIST.Row, 0) = "D"
        FILE_LIST.Rows(FILE_LIST.Row).StyleNew.ForeColor = Color.Gray

        For i = 1 To FILE_LIST.Rows.Count - 1
            If IsNumeric(FILE_LIST(i, 0)) = True Then
                FILE_LIST(i, 0) = i
            End If
        Next

    End Sub

    Private Sub BTN_REFRESH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_REFRESH.Click

        '최신 버전체크 파일을 다운로드 및 그리드에 표시
        FILE_LIST_LOAD()

    End Sub

    Private Sub BTN_SAVE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_SAVE.Click

        If FILE_LIST.Rows.Count = 1 Then
            MsgBox("내용이 없습니다.", MsgBoxStyle.Information, "YJ MMS : Program Upload")
            Exit Sub
        End If

        If MsgBox("저장 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "YJ MMS : Program Upload") = MsgBoxResult.No Then Exit Sub

        DBConnect_HO()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSql As String = String.Empty

        sqlTran = DBConnect2.BeginTransaction

        Try
            'FILE_LIST(0, 0) = "No"
            'FILE_LIST(0, 1) = "Division"
            'FILE_LIST(0, 2) = "File Name"
            'FILE_LIST(0, 3) = "Folder Name"
            'FILE_LIST(0, 4) = "Upload Date/Ttime"
            'FILE_LIST(0, 5) = "Upload ID"
            'FILE_LIST(0, 6) = "비고"
            For i = 1 To FILE_LIST.Rows.Count - 1
                If FILE_LIST(i, 0).ToString Like "N*" Then
                    strSql += "insert into PROGRAM_UPDATE(FILE_CODE, FILE_NAME, FOLDER_NAME, UPLOAD_DATE, WRITE_ID, UPLOAD_NOTE) values("
                    strSql += "'" & FILE_LIST(i, 1) & "'"
                    strSql += ",'" & FILE_LIST(i, 2) & "'"
                    strSql += ",'" & FILE_LIST(i, 3) & "'"
                    strSql += ",'" & FILE_LIST(i, 4) & "'"
                    strSql += ",'" & FILE_LIST(i, 5) & "'"
                    strSql += ",'" & FILE_LIST(i, 6) & "');"
                ElseIf FILE_LIST(i, 0).ToString Like "M*" Then
                    strSql += "update PROGRAM_UPDATE set FILE_NAME = '" & FILE_LIST(i, 2) & "'"
                    strSql += ", FOLDER_NAME = '" & FILE_LIST(i, 3) & "'"
                    strSql += ", UPLOAD_DATE = '" & Format(CDate(FILE_LIST(i, 4)), "yyyy-MM-dd HH:mm:ss") & "'"
                    strSql += ", WRITE_ID = '" & FILE_LIST(i, 5) & "'"
                    strSql += ", UPLOAD_NOTE = '" & FILE_LIST(i, 6) & "'"
                    strSql += " where FILE_CODE = '" & FILE_LIST(i, 1) & "';"
                ElseIf FILE_LIST(i, 0).ToString Like "D*" Then
                    strSql += "delete from PROGRAM_UPDATE where FILE_CODE = '" & FILE_LIST(i, 1) & "';"
                End If
                If FILE_LIST(i, 0).ToString Like "*,U" Then
                    Dim TEMP_FILE_Folder As String = Application.StartupPath & "\TEMP_FILE"
                    '폴더(생성)
                    If Not FILE_LIST(i, 3).ToString.ToUpper = "ROOT".ToUpper Then
                        'ROOT 폴더에서 Folde Name 폴더가 존재하는지 검색 후 존재하지 않으면 생성

                        Dim ftpFolder As String = ftpFolderSearch(ftpUrl & "/Run_File/")

                        Dim DirectoryName() As String = Split(ftpFolder, vbCrLf)
                        Dim DirectoryExists As Boolean

                        For j = 0 To UBound(DirectoryName) - 1
                            If FILE_LIST(i, 3) = Trim(DirectoryName(j)) Then
                                DirectoryExists = True
                                Exit For
                            Else
                                DirectoryExists = False
                            End If
                        Next

                        If DirectoryExists = False Then
                            ftpFolderMake(ftpUrl & "/Run_File/" & FILE_LIST(i, 3))
                        End If
                        '파일(업로드)-ROOT 이외
                        ftpFileUpload(ftpUrl & "/Run_File/" & FILE_LIST(i, 3), _
                                          TEMP_FILE_Folder & "\" & FILE_LIST(i, 2))
                    Else
                        '파일(업로드)-ROOT
                        ftpFileUpload(ftpUrl & "/Run_File", _
                                          TEMP_FILE_Folder & "\" & FILE_LIST(i, 2))
                    End If

                    '임시저장 파일 삭제
                    If My.Computer.FileSystem.FileExists(TEMP_FILE_Folder & "\" & FILE_LIST(i, 2)) Then _
                        My.Computer.FileSystem.DeleteFile(TEMP_FILE_Folder & "\" & FILE_LIST(i, 2))
                End If
            Next

            If strSql = String.Empty Then
                MsgBox("변경 사항이 없습니다.", MsgBoxStyle.Information, "YJ MMS : Program Upload")
                DBClose()
                Exit Sub
            End If

            sqlCmd = New MySqlCommand(strSql, DBConnect2)
            sqlCmd.Transaction = sqlTran
            sqlCmd.ExecuteNonQuery()

            sqlTran.Commit()
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, "YJ MMS : Program Upload")
            Exit Sub
        End Try

        DBClose_HO()

        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\NEWCheckVer.ini", True)
        file.WriteLine("[AUTO UPDATE]")
        For i = 1 To FILE_LIST.Rows.Count - 1
            file.WriteLine(FILE_LIST(i, 2) & "=" & FILE_LIST(i, 4) & "=" & FILE_LIST(i, 3))
        Next
        file.WriteLine("[/]")
        file.Close()

        '최신 버전체크파일 업로드
        ftpFileUpload(ftpUrl & "/Run_File", _
                          Application.StartupPath & "\NEWCheckVer.ini")
        '파일삭제
        System.IO.File.Delete(Application.StartupPath & "\NEWCheckVer.ini")

        MsgBox("저장완료.", MsgBoxStyle.Information, "YJ MMS : Program Upload")

        FILE_LIST_LOAD()

    End Sub
End Class
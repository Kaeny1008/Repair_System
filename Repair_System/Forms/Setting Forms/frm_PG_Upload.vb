Imports System.Threading
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_PG_Upload
    Private Sub frm_PG_Upload_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        GridSetting()

        'load_FileList()

    End Sub

    Private Sub GridSetting()

        With grid_FileList
            .AllowEditing = True
            .AllowFiltering = True
            .AllowSorting = AllowSortingEnum.SingleColumn
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 7
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_FileList(0, 0) = "No"
            grid_FileList(0, 1) = "Division"
            grid_FileList(0, 2) = "File Name"
            grid_FileList(0, 3) = "Folder Name"
            grid_FileList(0, 4) = "Upload Date/Ttime"
            grid_FileList(0, 5) = "Upload ID"
            grid_FileList(0, 6) = "비고"
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

    Private Sub load_FileList()

        grid_FileList.Rows.Count = 1

        DBConnect()

        Dim strSql As String = "select FILE_CODE, FILE_NAME, FOLDER_NAME, UPLOAD_DATE, WRITE_ID, ifnull(UPLOAD_NOTE, '') as UPLOAD_NOTE" &
            " from setting_program_manager order by FOLDER_NAME, FILE_CODE"

        Dim sqlCmd As New MySqlCommand(strSql, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_string As String = String.Empty
            insert_string = grid_FileList.Rows.Count & vbTab &
                            sqlDR("FILE_CODE") & vbTab &
                            sqlDR("FILE_NAME") & vbTab &
                            sqlDR("FOLDER_NAME") & vbTab &
                            Format(sqlDR("UPLOAD_DATE"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                            sqlDR("WRITE_ID") & vbTab &
                            sqlDR("UPLOAD_NOTE")
            grid_FileList.AddItem(insert_string)
        Loop
        sqlDR.Close()

        DBClose()

        grid_FileList.AutoSizeCols()

    End Sub

    Private Function FILE_LIST_BEFORE_LOAD(ByVal col_name As String, ByVal file_code As String) As String

        FILE_LIST_BEFORE_LOAD = String.Empty

        DBConnect()

        Dim strSql As String = "select ifnull(" & col_name & ", '') as " & col_name & " from setting_program_manager"
        strSql += " where file_code = '" & file_code & "'"

        Dim sqlCmd As New MySqlCommand(strSql, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            FILE_LIST_BEFORE_LOAD = sqlDR(col_name)
        Loop
        sqlDR.Close()

        DBClose()

        Return FILE_LIST_BEFORE_LOAD

    End Function

    Private Sub FILE_LIST_AfterEdit(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles grid_FileList.AfterEdit

        If IsNumeric(grid_FileList(e.Row, 0)) = False Then Exit Sub

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

        If FILE_LIST_BEFORE_LOAD(find_col_name, grid_FileList(e.Row, 1)) = grid_FileList(e.Row, e.Col) = False Then
            If Not grid_FileList(e.Row, 0).ToString = "N" Then
                grid_FileList.Rows(e.Row).StyleNew.ForeColor = Color.Red
                grid_FileList(e.Row, 0) = "M"
                grid_FileList(e.Row, 4) = Format(Now, "yyyy-MM-dd HH:mm:ss")
                grid_FileList(e.Row, 5) = loginID
            End If
        Else
            grid_FileList.Rows(e.Row).StyleNew.ForeColor = Color.Black
            grid_FileList(e.Row, 0) = "0"
            For i = 1 To grid_FileList.Rows.Count - 1
                If IsNumeric(grid_FileList(i, 0)) = True Then
                    grid_FileList(i, 0) = i
                End If
            Next
        End If

    End Sub

    Private Sub FILE_LIST_CellButtonClick(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles grid_FileList.CellButtonClick

        Dim openFile As New System.Windows.Forms.OpenFileDialog

        With openFile
            .Filter = "ALL Files (*.*)|*.*"
            .AddExtension() = True
        End With

        openFile.ShowDialog()

        If openFile.FileName = "" Then Exit Sub

        grid_FileList.Rows(e.Row).StyleNew.ForeColor = Color.Blue

        If IsNumeric(grid_FileList(e.Row, 0)) Then grid_FileList(e.Row, 0) = "M"

        Dim TEMP_FILE_Folder As String = Application.StartupPath & "\TEMP_FILE"

        '좌표폴더가 없으면 만들기
        If My.Computer.FileSystem.DirectoryExists(TEMP_FILE_Folder) = False Then
            My.Computer.FileSystem.CreateDirectory(TEMP_FILE_Folder)
        End If

        '선택한 파일을 임시폴더에 저장
        My.Computer.FileSystem.CopyFile(openFile.FileName, TEMP_FILE_Folder & "\" & openFile.FileName.Split("\")(UBound(openFile.FileName.Split("\"))), True)

        grid_FileList(e.Row, 2) = openFile.FileName.Split("\")(UBound(openFile.FileName.Split("\")))

        If Not grid_FileList(e.Row, 0) Like "*,U" Then
            grid_FileList(e.Row, 0) += ",U"
        End If

        grid_FileList(e.Row, 4) = Format(Now, "yyyy-MM-dd HH:mm:ss")
        grid_FileList(e.Row, 5) = loginID

        grid_FileList.AutoSizeCols()

    End Sub

    Private Sub FILE_LIST_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grid_FileList.MouseClick

        If e.Button = Windows.Forms.MouseButtons.Right And grid_FileList.MouseRow < 1 Then
            grid_FileList.Row = grid_FileList.MouseRow
            BTN_ADD.Enabled = False
            BTN_DELETE.Enabled = False
            cms_Grid.Show(grid_FileList, New Point(e.X, e.Y))
        ElseIf e.Button = Windows.Forms.MouseButtons.Right And grid_FileList.MouseRow > 0 Then
            grid_FileList.Row = grid_FileList.MouseRow
            BTN_ADD.Enabled = True
            BTN_DELETE.Enabled = True
            cms_Grid.Show(grid_FileList, New Point(e.X, e.Y))
        End If

    End Sub

    Private Sub FILE_LIST_RowColChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid_FileList.RowColChange

        If grid_FileList.Rows.Count = 1 Then Exit Sub
        If grid_FileList.Row < 1 Then Exit Sub

        Select Case grid_FileList.Col
            Case 1
                If IsNumeric(grid_FileList(grid_FileList.Row, 0)) Then
                    grid_FileList.AllowEditing = False
                Else
                    grid_FileList.AllowEditing = True
                End If
            Case Else
                grid_FileList.AllowEditing = True
        End Select

    End Sub

    Private Sub BTN_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_ADD.Click

        grid_FileList.AddItem("N" & vbTab & String.Empty & vbTab & String.Empty & vbTab & String.Empty & vbTab & Format(Now, "yyyy-MM-dd HH:mm:ss") & vbTab & loginID, grid_FileList.Row + 1)
        grid_FileList.TopRow = grid_FileList.Rows.Count - 1
        grid_FileList.Rows(grid_FileList.Row + 1).StyleNew.ForeColor = Color.Blue

        For i = 1 To grid_FileList.Rows.Count - 1
            If IsNumeric(grid_FileList(i, 0)) = True Then
                grid_FileList(i, 0) = i
            End If
        Next

    End Sub

    Private Sub BTN_DELETE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_DELETE.Click

        grid_FileList(grid_FileList.Row, 0) = "D"
        grid_FileList.Rows(grid_FileList.Row).StyleNew.ForeColor = Color.Gray

        For i = 1 To grid_FileList.Rows.Count - 1
            If IsNumeric(grid_FileList(i, 0)) = True Then
                grid_FileList(i, 0) = i
            End If
        Next

    End Sub

    Private Sub BTN_REFRESH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_REFRESH.Click

        thread_LoadingFormStart()
        '최신 버전체크 파일을 다운로드 및 그리드에 표시
        load_FileList()

        th_LoadingWindow.Abort()

    End Sub

    Private Sub BTN_SAVE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_SAVE.Click

        If grid_FileList.Rows.Count = 1 Then
            MsgBox("내용이 없습니다.", MsgBoxStyle.Information, "Repair System - 프로그램 파일 Upload")
            Exit Sub
        End If

        If MsgBox("저장 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Repair System - 프로그램 파일 Upload") = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Saving...")

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSql As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try
            For i = 1 To grid_FileList.Rows.Count - 1
                If grid_FileList(i, 0).ToString Like "N*" Then
                    strSql += "insert into setting_program_manager(FILE_CODE, FILE_NAME, FOLDER_NAME, UPLOAD_DATE, WRITE_ID, UPLOAD_NOTE) values("
                    strSql += "'" & grid_FileList(i, 1) & "'"
                    strSql += ",'" & grid_FileList(i, 2) & "'"
                    strSql += ",'" & grid_FileList(i, 3) & "'"
                    strSql += ",'" & grid_FileList(i, 4) & "'"
                    strSql += ",'" & grid_FileList(i, 5) & "'"
                    strSql += ",'" & grid_FileList(i, 6) & "');"
                ElseIf grid_FileList(i, 0).ToString Like "M*" Then
                    strSql += "update setting_program_manager set FILE_NAME = '" & grid_FileList(i, 2) & "'"
                    strSql += ", FOLDER_NAME = '" & grid_FileList(i, 3) & "'"
                    strSql += ", UPLOAD_DATE = '" & Format(CDate(grid_FileList(i, 4)), "yyyy-MM-dd HH:mm:ss") & "'"
                    strSql += ", WRITE_ID = '" & grid_FileList(i, 5) & "'"
                    strSql += ", UPLOAD_NOTE = '" & grid_FileList(i, 6) & "'"
                    strSql += " where FILE_CODE = '" & grid_FileList(i, 1) & "';"
                ElseIf grid_FileList(i, 0).ToString Like "D*" Then
                    strSql += "delete from setting_program_manager where FILE_CODE = '" & grid_FileList(i, 1) & "';"
                    ftpFileDelete(ftpUrl & "/Run_File",
                                          grid_FileList(i, 3), grid_FileList(i, 2))
                End If
                If grid_FileList(i, 0).ToString Like "*,U" Then
                    Dim TEMP_FILE_Folder As String = Application.StartupPath & "\TEMP_FILE"
                    '폴더(생성)
                    If Not grid_FileList(i, 3).ToString.ToUpper = "ROOT".ToUpper Then
                        'ROOT 폴더에서 Folde Name 폴더가 존재하는지 검색 후 존재하지 않으면 생성

                        Dim ftpFolder As String = ftpFolderSearch(ftpUrl & "/Run_File/")

                        Dim DirectoryName() As String = Split(ftpFolder, vbCrLf)
                        Dim DirectoryExists As Boolean

                        For j = 0 To UBound(DirectoryName) - 1
                            If grid_FileList(i, 3) = Trim(DirectoryName(j)) Then
                                DirectoryExists = True
                                Exit For
                            Else
                                DirectoryExists = False
                            End If
                        Next

                        If DirectoryExists = False Then
                            ftpFolderMake(ftpUrl & "/Run_File/" & grid_FileList(i, 3))
                        End If
                        '파일(업로드)-ROOT 이외
                        ftpFileUpload(ftpUrl & "/Run_File/" & grid_FileList(i, 3),
                                          TEMP_FILE_Folder & "\" & grid_FileList(i, 2))
                    Else
                        '파일(업로드)-ROOT
                        ftpFileUpload(ftpUrl & "/Run_File",
                                          TEMP_FILE_Folder & "\" & grid_FileList(i, 2))
                    End If

                    '임시저장 파일 삭제
                    If My.Computer.FileSystem.FileExists(TEMP_FILE_Folder & "\" & grid_FileList(i, 2)) Then _
                        My.Computer.FileSystem.DeleteFile(TEMP_FILE_Folder & "\" & grid_FileList(i, 2))
                End If
            Next

            If strSql = String.Empty Then
                'MsgBox("변경 사항이 없습니다.", MsgBoxStyle.Information, "Repair System - 프로그램 파일 Upload")
                th_LoadingWindow.Abort()
                Thread.Sleep(100)
                MessageBox.Show(Me, "변경 사항이 없습니다.", "Repair System - 프로그램 파일 Upload", MessageBoxButtons.OK, MessageBoxIcon.Information)
                DBClose()
                Exit Sub
            End If

            sqlCmd = New MySqlCommand(strSql, dbConnection1)
            sqlCmd.Transaction = sqlTran
            sqlCmd.ExecuteNonQuery()

            sqlTran.Commit()
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Repair System - 프로그램 파일 Upload")
            Exit Sub
        End Try

        DBClose()

        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\NEWCheckVer.ini", True)
        file.WriteLine("[AUTO UPDATE]")
        For i = 1 To grid_FileList.Rows.Count - 1
            If Not grid_FileList(i, 0).ToString Like "D*" Then
                file.WriteLine(grid_FileList(i, 2) & "=" & grid_FileList(i, 4) & "=" & grid_FileList(i, 3))
            End If
        Next
        file.WriteLine("[/]")
        file.Close()

        '최신 버전체크파일 업로드
        ftpFileUpload(ftpUrl & "/Run_File",
                          Application.StartupPath & "\NEWCheckVer.ini")
        '파일삭제
        System.IO.File.Delete(Application.StartupPath & "\NEWCheckVer.ini")

        load_FileList()

        th_LoadingWindow.Abort()
        Thread.Sleep(100)
        MessageBox.Show(Me, "저장완료.", "Repair System - 프로그램 파일 Upload", MessageBoxButtons.OK, MessageBoxIcon.Information)


    End Sub

    Private Sub grid_FileList_Click(sender As Object, e As EventArgs) Handles grid_FileList.Click

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        thread_LoadingFormStart()
        '최신 버전체크 파일을 다운로드 및 그리드에 표시
        load_FileList()

        th_LoadingWindow.Abort()

    End Sub
End Class
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_code_manager
    Private Sub frm_CodeManager_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        GridSetting()

    End Sub

    Private Sub GridSetting()

        With grid_DivisionList
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 5
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_DivisionList(0, 0) = "No"
            grid_DivisionList(0, 1) = "분류"
            grid_DivisionList(0, 2) = "CODE"
            grid_DivisionList(0, 3) = "CODE Name"
            grid_DivisionList(0, 4) = "비고"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(3).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

        With grid_DetailList
            .AllowEditing = True
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 4
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_DetailList(0, 0) = "No"
            grid_DetailList(0, 1) = "CODE"
            grid_DetailList(0, 2) = "CODE Name"
            grid_DetailList(0, 3) = "비고"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(3).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ExtendLastCol = True
            .ShowCursor = True
        End With

    End Sub

    Private Sub DivisionLoad()

        DBConnect()

        grid_DivisionList.Rows.Count = 1

        Dim strSql As String = String.Empty

        strSql = "select * from setting_code_manager order by division_name"

        Dim sqlCmd As New MySqlCommand(strSql, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_DivisionList.Rows.Count & vbTab &
                                          sqlDR("division_high_rank") & vbTab &
                                          sqlDR("division_code") & vbTab &
                                          sqlDR("division_name") & vbTab &
                                          sqlDR("division_note")

            grid_DivisionList.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        grid_DivisionList.AutoSizeCols()

    End Sub

    Dim highrank_CODE As String

    Private Sub grid_DivisionList_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grid_DivisionList.MouseDoubleClick

        If e.Button = Windows.Forms.MouseButtons.Left And grid_DivisionList.MouseRow > 0 Then
            highrank_CODE = grid_DivisionList(grid_DivisionList.MouseRow, 2) '상위코드 변수 지정
            LowrankLOAD()
        End If

    End Sub

    Private Sub LowrankLOAD()

        DBConnect()

        grid_DetailList.Rows.Count = 1

        Dim strSql As String = String.Empty

        strSql = "select * from setting_code_manager_detail where division_code = '" & highrank_CODE & "' order by detail_code"

        Dim sqlCmd As New MySqlCommand(strSql, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_DetailList.Rows.Count & vbTab &
                                          sqlDR("detail_code") & vbTab &
                                          sqlDR("detail_name") & vbTab &
                                          sqlDR("detail_note")

            grid_DetailList.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        If grid_DetailList.Rows.Count = 1 Then
            grid_DetailList.AddItem("N")
            grid_DetailList.Rows(grid_DetailList.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
        End If

        grid_DetailList.AutoSizeCols()

    End Sub

    Private Sub btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add.Click

        grid_DetailList.AddItem("N")
        grid_DetailList.TopRow = grid_DetailList.Rows.Count - 1
        grid_DetailList.Rows(grid_DetailList.Rows.Count - 1).StyleNew.ForeColor = Color.Blue

    End Sub

    Private Sub grid_DetailListAfterEdit(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles grid_DetailList.AfterEdit

        If before_Text = grid_DetailList(e.Row, e.Col) Then Exit Sub

        '저장된 CODE를 수정할 경우 수정할 수 없다고 알린다.
        If IsNumeric(grid_DetailList(e.Row, 0)) And e.Col = 1 Then
            grid_DetailList(e.Row, e.Col) = before_Text
            MsgBox("저장된 CODE는 수정할 수 없습니다.", MsgBoxStyle.Information, "Repair System - Code 관리")
            Exit Sub
        End If

        grid_DetailList.AutoSizeCols()

        If IsNumeric(grid_DetailList(e.Row, 0)) = True Then
            grid_DetailList(e.Row, 0) = "M"
            grid_DetailList.Rows(e.Row).StyleNew.ForeColor = Color.Red
        End If


        If e.Col = 1 Then

            If Not grid_DetailList(e.Row, e.Col).ToString.Length = 5 Then
                MsgBox("CODE는 5자리입니다.", MsgBoxStyle.Information, "Repair System - Code 관리")
                grid_DetailList(e.Row, e.Col) = String.Empty
                Exit Sub
            End If

            grid_DetailList(e.Row, e.Col) = grid_DetailList(e.Row, e.Col).ToString.ToUpper

            Dim edit_TEXT As String = grid_DetailList(e.Row, e.Col)
            Dim start_CODE As String = grid_DetailList(1, 1).ToString.Substring(0, 2)
            For i = 1 To grid_DetailList.Rows.Count - 1
                If Not grid_DetailList(i, 1) = String.Empty Then
                    If Not start_CODE = grid_DetailList(i, 1).ToString.Substring(0, 2) Then
                        MsgBox("시작 CODE가 다릅니다.", MsgBoxStyle.Information, "Repair System - Code 관리")
                        grid_DetailList(e.Row, e.Col) = String.Empty
                        Exit Sub
                    End If
                    If Not i = e.Row Then
                        If edit_TEXT = grid_DetailList(i, 1) Then
                            MsgBox("중복 CODE가 있습니다.", MsgBoxStyle.Information, "Repair System - Code 관리")
                            grid_DetailList(e.Row, e.Col) = String.Empty
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End If

    End Sub

    Dim before_Text As String

    Private Sub DETAIL_LIST_BeforeEdit(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles grid_DetailList.BeforeEdit

        before_Text = grid_DetailList(e.Row, e.Col)

    End Sub

    Private Sub detail_list_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grid_DetailList.KeyDown

        If e.KeyCode = 93 Then
            If grid_DetailList.Row > 0 And grid_DetailList.Col > 0 Then
                cms_Grid.Show(grid_DetailList, New Point(MousePosition.X - 420, MousePosition.Y - 53))
            End If
        End If

    End Sub

    Private Sub DETAIL_LIST_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grid_DetailList.MouseClick

        If e.Button = Windows.Forms.MouseButtons.Right And grid_DetailList.MouseRow > 0 Then
            grid_DetailList.Row = grid_DetailList.MouseRow
            cms_Grid.Show(grid_DetailList, New Point(e.X, e.Y))
        End If

    End Sub

    Private Sub BTN_DELETE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Delete.Click

        grid_DetailList(grid_DetailList.Row, 0) = "D"
        grid_DetailList.Rows(grid_DetailList.Row).StyleNew.ForeColor = Color.Gray

    End Sub

    Private Sub BTN_SAVE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save.Click

        If MsgBox("저장 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Repair System - Code 관리") = MsgBoxResult.No Then Exit Sub

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try
            '누락 부분 검사 및 sql작성
            For i = 1 To grid_DetailList.Rows.Count - 1
                For j = 1 To grid_DetailList.Cols.Count - 2
                    If grid_DetailList(i, j) = String.Empty Then
                        Dim col_String As String = String.Empty
                        If j = 1 Then
                            col_String = "CODE"
                        ElseIf j = 2 Then
                            col_String = "CODE Name"
                        End If
                        MsgBox("행 : " & i & vbCrLf &
                               col_String & "이(가) 입력되지 않았습니다.", MsgBoxStyle.Information, "Repair System - Code 관리")
                        Exit Sub
                    End If
                Next
                '------------- 여기에 sql 작성
                If grid_DetailList(i, 0) = "D" Then '삭제라면
                    strSQL += "delete from setting_code_manager_detail where detail_code = '" & grid_DetailList(i, 1) & "';"
                ElseIf grid_DetailList(i, 0) = "M" Then '수정이라면
                    strSQL += "update setting_code_manager_detail set detail_name = '" & grid_DetailList(i, 2) & "'"
                    strSQL += ", detail_note = '" & grid_DetailList(i, 3) & "'"
                    strSQL += " where detail_code = '" & grid_DetailList(i, 1) & "';"
                ElseIf grid_DetailList(i, 0) = "N" Then '신규라면
                    strSQL += "insert into setting_code_manager_detail values("
                    strSQL += "'" & grid_DetailList(i, 1) & "'"
                    strSQL += ",'" & grid_DetailList(i, 2) & "'"
                    strSQL += ",'" & grid_DetailList(i, 3) & "'"
                    strSQL += ",'" & highrank_CODE & "');"
                End If
            Next

            If strSQL = String.Empty Then
                MsgBox("변경사항이 없습니다.", MsgBoxStyle.Information, "Repair System - Code 관리")
                DBClose()
                LowrankLOAD()
                Exit Sub
            End If

            sqlCmd = New MySqlCommand(strSQL, dbConnection1)
            sqlCmd.Transaction = sqlTran
            sqlCmd.ExecuteNonQuery()

            sqlTran.Commit()
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Repair System - Code 관리")
            Exit Sub
        End Try

        DBClose()

        MsgBox("저장완료.", MsgBoxStyle.Information, "Repair System - Code 관리")

        LowrankLOAD()

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        DivisionLoad()

    End Sub
End Class
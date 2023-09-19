'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-08-08
'##### 수정일자 : 2023-08-08
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 

'############################################################################################################
'############################################################################################################

Imports System.Threading
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_Reject_IC_In

    Dim form_Msgbox_String As String = "폐기자재 회수"

    Private Sub frm_Reject_IC_In_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        grid_Setting()

        DateTimePicker1.Value = Format(Now, "yyyy-MM-01 00:00:00")
        DateTimePicker2.Value = Format(Now, "yyyy-MM-dd 23:59:59")

    End Sub

    Private Sub grid_Setting()

        With grid_RejectList
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 4
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_RejectList(0, 0) = "No"
            grid_RejectList(0, 1) = "회수 일자"
            grid_RejectList(0, 2) = "회수번호"
            grid_RejectList(0, 3) = "Lot 수"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

        With grid_RejectList_Detail
            .AllowEditing = False
            .AllowFiltering = True
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 12
            .Cols.Fixed = 1
            .Rows.Count = 2
            .Rows.Fixed = 2
            .AutoClipboard = True
            .AllowFreezing = AllowFreezingEnum.Both
            .AllowMerging = AllowMergingEnum.FixedOnly
            .AllowMergingFixed = AllowMergingEnum.FixedOnly
            For i = 0 To .Cols.Count - 1
                .Cols(i).AllowMerging = True
            Next
            .Rows(0).AllowMerging = True
            .Rows(1).AllowMerging = True
            Dim Mrng As CellRange = .GetCellRange(0, 0, 1, 0)
            Mrng.Data = "No"
            Mrng = .GetCellRange(0, 1, 1, 1)
            Mrng.Data = "YJ No"
            Mrng = .GetCellRange(0, 2, 1, 4)
            Mrng.Data = "Module"
            grid_RejectList_Detail(1, 2) = "Product Code"
            grid_RejectList_Detail(1, 3) = "Lot No."
            grid_RejectList_Detail(1, 4) = "수량(EA)"
            Mrng = .GetCellRange(0, 5, 1, 10)
            Mrng.Data = "회수정보"
            grid_RejectList_Detail(1, 5) = "PMIC"
            grid_RejectList_Detail(1, 6) = "회수수량"
            grid_RejectList_Detail(1, 7) = "확인"
            grid_RejectList_Detail(1, 8) = "RCD"
            grid_RejectList_Detail(1, 9) = "회수수량"
            grid_RejectList_Detail(1, 10) = "확인"
            .Cols(7).ComboList = "OK|NG"
            .Cols(10).ComboList = "OK|NG"
            Mrng = .GetCellRange(0, 11, 1, 11)
            Mrng.Data = "비고"
            '.Styles("Normal").WordWrap = True
            '.Cols(11).StyleNew.WordWrap = True

            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 1).TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
            .ExtendLastCol = True
        End With

    End Sub

    Private Sub grid_RejectList_Detail_RowColChange(sender As Object, e As EventArgs) Handles grid_RejectList_Detail.RowColChange

        Select Case grid_RejectList_Detail.Col
            Case 5, 6, 7, 8, 9, 10, 11
                grid_RejectList_Detail.AllowEditing = True
            Case Else
                grid_RejectList_Detail.AllowEditing = False
        End Select

    End Sub

    Private Sub btn_New_Reject_Click(sender As Object, e As EventArgs) Handles btn_New_Reject.Click

        now_Total_Lot = 0
        now_Total_Module = 0
        RN_Code_Maker()
        tb_Barcode.Text = String.Empty
        tb_Barcode.Enabled = True
        tb_Barcode.SelectAll()
        tb_Barcode.Focus()
        btn_Save.Enabled = True

        grid_RejectList_Detail.Rows.Count = 2

    End Sub

    Private Sub RN_Code_Maker()

        Dim rn_code As String = String.Empty

        DBConnect()

        Try
            Dim strSQL As String = "Call sp_reject_material_collect(0" &
                        ", '" & Format(Now, "yyyy-MM-dd") & "'" &
                        ", null" &
                        ", null" &
                        ", null)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                If Not IsDBNull(sqlDR("reject_no")) Then rn_code = "RN" & Format(Now, "yyMMdd") & sqlDR("reject_no")
            Loop
            sqlDR.Close()
        Catch ex As Exception
            MsgBox("회수번호 생성중 에러 발생 : " & vbCrLf & ex.Message,
                   MsgBoxStyle.Critical,
                   form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        If rn_code = String.Empty Then rn_code = "RN" & Format(Now, "yyMMdd") & "001"

        TextBox1.Text = rn_code

    End Sub

    Private Sub tb_Barcode_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_Barcode.KeyDown

        If e.KeyCode = 13 And Not Trim(tb_Barcode.Text) = String.Empty Then
            thread_LoadingFormStart()
            tb_Barcode.Text = tb_Barcode.Text.ToUpper

            If UBound(tb_Barcode.Text.Split("/")) < 2 And Not CheckBox1.Checked Then
                MsgBox("Yujin Barcode(Matamatrix)를 스캔하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
                tb_Barcode.SelectAll()
                tb_Barcode.Focus()
                Exit Sub
            End If

            Dim use_lot_no As String = tb_Barcode.Text.Split("/")(0)
            If CheckBox1.Checked Then use_lot_no = tb_Barcode.Text

            Dim findIndex As Integer = 3
            If CheckBox1.Checked Then findIndex = 1

            For i = 2 To grid_RejectList_Detail.Rows.Count - 1
                If use_lot_no.Equals(grid_RejectList_Detail(i, findIndex)) Then
                    MsgBox("이미 등록된 Lot 입니다.",
                           MsgBoxStyle.Information,
                           form_Msgbox_String)
                    tb_Barcode.SelectAll()
                    tb_Barcode.Focus()
                    Exit Sub
                End If
            Next

            th_LoadingWindow.Abort()

            Dim status_check As String = lot_Information_Load(use_lot_no)
            If status_check = String.Empty Then 'Lot 정보를 불러온다.
                MsgBox("현장 출고 중인 Lot이 아닙니다.",
                       MsgBoxStyle.Information,
                       form_Msgbox_String)
            ElseIf status_check = "Completed" Then 'Lot 정보를 불러온다.
                MsgBox("이미 회수 완료된 Lot 입니다.",
                       MsgBoxStyle.Information,
                       form_Msgbox_String)
            End If
        End If

    End Sub

    Dim now_Total_Module As Integer = 0
    Dim now_Total_Lot As Integer = 0

    Private Function lot_Information_Load(ByVal lot_no As String) As String

        lot_Information_Load = String.Empty

        grid_RejectList_Detail.Redraw = False

        DBConnect()

        Try
            Dim strSQL As String = "call sp_reject_material_collect(1" &
            ", null" &
            ", null" &
            ", '" & lot_no & "'" &
            ", null)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                Dim pmic_part_no As String = sqlDR("pmic_part_no")
                Dim pmic_qty As String = Format(sqlDR("chip_qty"), "#,##0")
                Dim rcd_part_no As String = sqlDR("rcd_part_no")
                Dim rcd_qty As String = Format(sqlDR("chip_qty"), "#,##0")

                If sqlDR("lot_option") Like "*O" Then
                    rcd_part_no = "None"
                    rcd_qty = 0
                End If

                Dim insert_string As String = "N" & vbTab &
                                              sqlDR("yj_no") & vbTab &
                                              sqlDR("product") & vbTab &
                                              sqlDR("lot_no") & vbTab &
                                              Format(sqlDR("chip_qty"), "#,##0") & vbTab &
                                              pmic_part_no & vbTab &
                                              pmic_qty & vbTab &
                                              "OK" & vbTab &
                                              rcd_part_no & vbTab &
                                              rcd_qty & vbTab &
                                              "OK"
                If sqlDR("lot_status").ToString.Equals("Moving to WorkSite") Or
                    sqlDR("lot_status").ToString.Equals("Completed") Or
                    sqlDR("lot_status").ToString.Equals("SMT PMIC Working") Or
                    sqlDR("lot_status").ToString.Equals("PMIC Working Completed") Or
                    sqlDR("lot_status").ToString.Equals("SMT RCD Working") Or
                    sqlDR("lot_status").ToString.Equals("RCD Working Completed") Then
                    If Not IsDBNull(sqlDR("reject_material")) Then
                        lot_Information_Load = "Completed"
                    Else
                        lot_Information_Load = "OK"
                        grid_RejectList_Detail.AddItem(insert_string)
                        grid_RejectList_Detail.Rows(grid_RejectList_Detail.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
                        now_Total_Module += sqlDR("chip_qty")
                        now_Total_Lot += 1
                    End If
                End If
            Loop
            sqlDR.Close()
        Catch ex As Exception
            MsgBox("Lot 정보 불러오는 중 에러 발생 : " & ex.Message,
                    MsgBoxStyle.Information,
                    form_Msgbox_String)
        End Try

        DBClose()

        Label10.Text = "총 회수 Lot수 : " & Format(now_Total_Lot, "#,##0") & " Lot"
        Label3.Text = "총 회수 Module : " & Format(now_Total_Module, "#,##0") & " EA"

        grid_RejectList_Detail.AutoSizeCols()
        grid_RejectList_Detail.AutoSizeRows(2, 0, grid_RejectList_Detail.Rows.Count - 1, grid_RejectList_Detail.Cols.Count - 1, 0, AutoSizeFlags.None)
        grid_RejectList_Detail.Redraw = True

        Return lot_Information_Load

    End Function

    Private Sub btn_Search_Click(sender As Object, e As EventArgs) Handles btn_Search.Click

        thread_LoadingFormStart()

        tb_Barcode.Text = "스캔 불가시 'Lot No/YJ No/'로 입력하여 주세요."
        TextBox1.Text = String.Empty
        tb_Barcode.Enabled = False

        grid_RejectList_Detail.Rows.Count = 2
        grid_RejectList.Rows.Count = 1
        grid_RejectList.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_reject_material_collect(2" &
                                ",'" & Format(DateTimePicker1.Value, "yyyy-MM-dd 00:00:00") & "'" &
                                ",'" & Format(DateTimePicker2.Value, "yyyy-MM-dd 23:59:59") & "'" &
                                ",'" & tb_S_Lot_No.Text & "'" &
                                ",'" & tb_S_YJ_No.Text & "')"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_RejectList.Rows.Count & vbTab &
                                          Format(sqlDR("reject_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          sqlDR("reject_no") & vbTab &
                                          Format(sqlDR("reject_lot_qty"), "#,##0")
            grid_RejectList.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        grid_RejectList.AutoSizeCols()
        grid_RejectList.Redraw = True
        now_Total_Lot = 0
        now_Total_Module = 0

        th_LoadingWindow.Abort()

    End Sub

    Private Sub tb_Barcode_TextChanged(sender As Object, e As EventArgs) Handles tb_Barcode.TextChanged

    End Sub

    Private Sub btn_Save_Click(sender As Object, e As EventArgs) Handles btn_Save.Click

        If grid_RejectList_Detail.Rows.Count = 2 Then
            MsgBox("저장할 내용이 없습니다.",
                   MsgBoxStyle.Information,
                   form_Msgbox_String)
            Exit Sub
        End If

        If MsgBox("저장 하시겠습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                  form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Saving...")

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try
            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            '라벨 발행 기록
            For i = 1 To grid_RejectList_Detail.Rows.Count - 1
                If grid_RejectList_Detail(i, 7) = "NG" Or
                    grid_RejectList_Detail(i, 10) = "NG" Then
                    MsgBox("확인 결과 NG인 항목이 있습니다." & vbCrLf & "행 : " & i - 1,
                           MsgBoxStyle.Information,
                           form_Msgbox_String)
                    Exit Sub
                End If
                Select Case grid_RejectList_Detail(i, 0).ToString
                    Case "N"
                        strSQL += "insert into reject_material_collect(reject_no, lot_no, yj_no"
                        strSQL += ", pmic_qty, pmic_result, rcd_qty, rcd_result, reject_note"
                        strSQL += ", reject_date, writer, write_date, first_write_date) values("
                        strSQL += "'" & TextBox1.Text & "'"
                        strSQL += ",'" & grid_RejectList_Detail(i, 3) & "'"
                        strSQL += ",'" & grid_RejectList_Detail(i, 1) & "'"
                        strSQL += ",'" & grid_RejectList_Detail(i, 6) & "'"
                        strSQL += ",'" & grid_RejectList_Detail(i, 7) & "'"
                        strSQL += ",'" & grid_RejectList_Detail(i, 9) & "'"
                        strSQL += ",'" & grid_RejectList_Detail(i, 10) & "'"
                        strSQL += ",'" & grid_RejectList_Detail(i, 11) & "'"
                        strSQL += ",'" & write_date & "'"
                        strSQL += ",'" & loginID & "'"
                        strSQL += ",'" & write_date & "'"
                        strSQL += ",'" & write_date & "');"

                        strSQL += "update basic_lot_information set reject_material = 'OK'"
                        strSQL += " where lot_no = '" & grid_RejectList_Detail(i, 3) & "';"
                    Case "M"
                        strSQL += "update reject_material_collect set"
                        strSQL += " reject_note = '" & grid_RejectList_Detail(i, 11) & "'"
                        strSQL += ", writer = '" & loginID & "'"
                        strSQL += ", write_date = '" & write_date & "'"
                        strSQL += " where reject_no = '" & TextBox1.Text & "'"
                        strSQL += " and lot_no = '" & grid_RejectList_Detail(i, 3) & "';"
                    Case "D"
                        strSQL += "update basic_lot_information set reject_material = null"
                        strSQL += " where lot_no = '" & grid_RejectList_Detail(i, 3) & "';"
                        strSQL += "delete from reject_material_collect"
                        strSQL += " where reject_no = '" & TextBox1.Text & "'"
                        strSQL += " and lot_no = '" & grid_RejectList_Detail(i, 3) & "'"
                End Select
            Next

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        th_LoadingWindow.Abort()
        Thread.Sleep(100)
        MsgBox("폐기자재 회수 내용을 저장 하였습니다.", MsgBoxStyle.Information, form_Msgbox_String)

        grid_RejectList_Detail.Rows.Count = 2
        btn_Save.Enabled = False
        TextBox1.Text = String.Empty
        tb_Barcode.Text = String.Empty
        now_Total_Lot = 0
        now_Total_Module = 0

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub grid_RejectList_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles grid_RejectList.MouseDoubleClick

        Dim sel_row As Integer = grid_RejectList.MouseRow

        If e.Button = MouseButtons.Left And sel_row > 0 Then

            thread_LoadingFormStart()

            grid_RejectList_Detail.Rows.Count = 2
            grid_RejectList_Detail.Redraw = False

            TextBox1.Text = grid_RejectList(sel_row, 2)
            btn_Save.Enabled = True
            tb_Barcode.Enabled = True
            tb_Barcode.Text = String.Empty
            tb_Barcode.SelectAll()
            tb_Barcode.Focus()

            DBConnect()

            Dim total_Lot As Integer = 0
            Dim total_Module As Integer = 0

            Dim strSQL As String = "call sp_reject_material_collect(3" &
                                    ",null" &
                                    ",null" &
                                    ",'" & grid_RejectList(sel_row, 2) & "'" &
                                    ",null)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                Dim insert_String As String = grid_RejectList_Detail.Rows.Count - 1 & vbTab &
                                              sqlDR("yj_no") & vbTab &
                                              sqlDR("product") & vbTab &
                                              sqlDR("lot_no") & vbTab &
                                              sqlDR("chip_qty") & vbTab &
                                              sqlDR("pmic_part_no") & vbTab &
                                              sqlDR("pmic_qty") & vbTab &
                                              sqlDR("pmic_result") & vbTab &
                                              sqlDR("rcd_part_no") & vbTab &
                                              sqlDR("rcd_qty") & vbTab &
                                              sqlDR("rcd_result") & vbTab &
                                              sqlDR("reject_note")
                grid_RejectList_Detail.AddItem(insert_String)
                total_Lot += 1
                total_Module += sqlDR("chip_qty")
            Loop
            sqlDR.Close()

            DBClose()

            Label10.Text = "총 회수 Lot수 : " & Format(total_Lot, "#,##0") & " Lot"
            Label3.Text = "총 회수 Module : " & Format(total_Module, "#,##0") & " EA"

            grid_RejectList_Detail.AutoSizeCols()
            grid_RejectList_Detail.Redraw = True

            th_LoadingWindow.Abort()
        End If

    End Sub

    Private Sub grid_RejectList_Detail_AfterEdit(sender As Object, e As RowColEventArgs) Handles grid_RejectList_Detail.AfterEdit

        If IsNumeric(grid_RejectList_Detail(e.Row, 0)) Then
            grid_RejectList_Detail(e.Row, 0) = "M"
            grid_RejectList_Detail.Rows(e.Row).StyleNew.ForeColor = Color.Red
        End If

        grid_RejectList_Detail.Redraw = False
        grid_RejectList_Detail.AutoSizeCols()
        grid_RejectList_Detail.Redraw = True

    End Sub

    Private Sub grid_RejectList_Detail_MouseClick(sender As Object, e As MouseEventArgs) Handles grid_RejectList_Detail.MouseClick

        Dim sel_row As Integer = grid_RejectList_Detail.MouseRow

        If sel_row < 2 Then Exit Sub
        grid_RejectList_Detail.Row = sel_row

        If e.Button = MouseButtons.Right Then
            cms_Menu.Show(grid_RejectList_Detail, New Point(e.X, e.Y))
        End If

    End Sub

    Private Sub btn_RowDelete_Click(sender As Object, e As EventArgs) Handles btn_RowDelete.Click

        If grid_RejectList_Detail(grid_RejectList_Detail.Row, 0).ToString = "N" Then
            grid_RejectList_Detail.RemoveItem(grid_RejectList_Detail.Row)
        Else
            grid_RejectList_Detail(grid_RejectList_Detail.Row, 0) = "D"
        End If

        grid_RejectList_Detail.Redraw = False

        For i = 1 To grid_RejectList_Detail.Rows.Count - 1
            If IsNumeric(grid_RejectList_Detail(i, 0)) Then
                grid_RejectList_Detail(i, 0) = i
                grid_RejectList_Detail.Rows(i).StyleNew.ForeColor = Color.Black
            Else
                If grid_RejectList_Detail(i, 0).ToString = "N" Then
                    grid_RejectList_Detail.Rows(i).StyleNew.ForeColor = Color.Blue
                ElseIf grid_RejectList_Detail(i, 0).ToString = "D" Then
                    grid_RejectList_Detail.Rows(i).StyleNew.ForeColor = Color.Gray
                End If
            End If
        Next

        grid_RejectList_Detail.AutoSizeCols()

        grid_RejectList_Detail.Redraw = True

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        tb_Barcode.SelectAll()
        tb_Barcode.Focus()

    End Sub

    Private Sub grid_RejectList_MouseMove(sender As Object, e As MouseEventArgs) Handles grid_RejectList.MouseMove, grid_RejectList_Detail.MouseMove

        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then

            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub

    Private Sub grid_LotList_LostFocus(sender As Object, e As EventArgs) Handles grid_RejectList.LostFocus,
                                                                                 grid_RejectList_Detail.LostFocus

        frm_Main.lb_Status.Text = String.Empty

    End Sub
End Class
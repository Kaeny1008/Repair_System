'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-07-24
'##### 수정일자 : 2023-07-24
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 베이킹 작업 이력을 확인 한다.

'############################################################################################################
'############################################################################################################

Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_Baking_History
    Private Sub frm_Baking_History_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        grid_Setting()

        Start_Date.Value = Format(Now, "yyyy-MM-01 00:00:00")
        End_Date.Value = Format(Now, "yyyy-MM-dd HH:mm:ss")

    End Sub

    Private Sub grid_Setting()

        With grid_DailyList
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 3
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_DailyList(0, 0) = "No"
            grid_DailyList(0, 1) = "작업일자"
            grid_DailyList(0, 2) = "작업수량"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .AutoSizeCols()
            .ShowCursor = True
            .Tree.Style = TreeStyleFlags.Simple
        End With

        With grid_LotList
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 13
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_LotList(0, 0) = "No"
            grid_LotList(0, 1) = "구분"
            grid_LotList(0, 2) = "시작일자"
            grid_LotList(0, 3) = "완료일자"
            grid_LotList(0, 4) = "작업시간(분)"
            grid_LotList(0, 5) = "Product Code"
            grid_LotList(0, 6) = "Lot No."
            grid_LotList(0, 7) = "Lot Option"
            grid_LotList(0, 8) = "YJ No."
            grid_LotList(0, 9) = "Qty"
            grid_LotList(0, 10) = "작업자"
            grid_LotList(0, 11) = "작업호기"
            grid_LotList(0, 12) = "비고"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .ExtendLastCol = True
            .AutoSizeCols()
            .ShowCursor = True
            .Tree.Style = TreeStyleFlags.Simple
        End With

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        thread_LoadingFormStart()

        grid_DailyList.Redraw = False
        grid_DailyList.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call sp_baking_history(0" &
            ", '" & Format(Start_Date.Value, "yyyy-MM-dd 00:00:00") & "'" &
            ",'" & Format(End_Date.Value, "yyyy-MM-dd 23:59:59") & "'" &
            ", '" & tb_Lot_No.Text.ToUpper & "'" &
            ", null)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_DailyList.Rows.Count & vbTab &
                                          sqlDR("working_date") & vbTab &
                                          Format(sqlDR("working_count"), "#,##0")
            grid_DailyList.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        'Dim s As CellStyle = c1flexgrid1.Styles(CellStyleEnum.Subtotal0)
        's.BackColor = Color.AliceBlue
        's.ForeColor = Color.Black
        's.Format = "#,##0"
        'c1flexgrid1.Subtotal(AggregateEnum.Sum, 0, 1, 2, "구분합계")

        grid_DailyList.AutoSizeCols()
        grid_DailyList.Redraw = True

        th_LoadingWindow.Abort()

    End Sub

    Private Sub grid_DailyList_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles grid_DailyList.MouseDoubleClick

        If e.Button = MouseButtons.Left And grid_DailyList.MouseRow > 0 Then
            Dim selectDate As String = grid_DailyList(grid_DailyList.MouseRow, 1)
            dailyList_Load(selectDate)
        End If

    End Sub

    Private Sub dailyList_Load(ByVal selectDate As String)

        thread_LoadingFormStart()

        grid_LotList.Redraw = False
        grid_LotList.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call sp_baking_history(2, null, null, '" & selectDate & "', null)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_LotList.Rows.Count & vbTab &
                                          sqlDR("part_division") & vbTab &
                                          Format(sqlDR("start_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          Format(sqlDR("end_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          DateDiff(DateInterval.Minute, sqlDR("start_date"), sqlDR("end_date")) & vbTab &
                                          sqlDR("product") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("lot_option") & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          Format(sqlDR("working_count"), "#,##0") & vbTab &
                                          sqlDR("baking_worker") & vbTab &
                                          sqlDR("machine_no")
            grid_LotList.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        Dim s As CellStyle = grid_LotList.Styles(CellStyleEnum.Subtotal0)
        s.BackColor = Color.AliceBlue
        s.ForeColor = Color.Black
        s.Format = "#,##0"
        grid_LotList.SubtotalPosition = SubtotalPositionEnum.BelowData
        grid_LotList.Subtotal(AggregateEnum.Sum, 0, 1, 9, "구분합계")

        grid_LotList.AutoSizeCols()
        grid_LotList.Redraw = True

        th_LoadingWindow.Abort()

    End Sub

    Private Sub grid_LotList_MouseMove(sender As Object, e As MouseEventArgs) Handles grid_LotList.MouseMove,
                                                                                      grid_DailyList.MouseMove

        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then

            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub

    Private Sub grid_LotList_LostFocus(sender As Object, e As EventArgs) Handles grid_LotList.LostFocus,
                                                                                 grid_DailyList.LostFocus

        frm_Main.lb_Status.Text = String.Empty

    End Sub

    Private Sub grid_LotList_KeyDown(sender As Object, e As KeyEventArgs) Handles grid_LotList.KeyDown

        If e.Shift Then
            md_ETC.cellCal(sender, Nothing)
        Else
            frm_Main.lb_Status.Text = String.Empty
        End If

    End Sub
End Class
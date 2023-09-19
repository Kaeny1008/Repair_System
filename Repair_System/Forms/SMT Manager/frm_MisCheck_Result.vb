Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_MisCheck_Result
    Private Sub frm_MisCheck_Result_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Start_Date.Value = Format(Now, "yyyy-MM-dd 00:00:00")
        End_Date.Value = Format(Now, "yyyy-MM-dd 23:59:59")

        GRID_SETTING()

    End Sub

    Private Sub GRID_SETTING()

        With grid_LotList
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 6
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_LotList(0, 0) = "No"
            grid_LotList(0, 1) = "Lot No."
            grid_LotList(0, 2) = "Option"
            grid_LotList(0, 3) = "YJ No."
            grid_LotList(0, 4) = "작업구분"
            grid_LotList(0, 5) = "작업일자"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

        With grid_PartHistory
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 14
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_PartHistory(0, 0) = "No"
            grid_PartHistory(0, 1) = "기준 Part No."
            grid_PartHistory(0, 2) = "사용 Part No."
            grid_PartHistory(0, 3) = "Lot No."
            grid_PartHistory(0, 4) = "기초재고"
            grid_PartHistory(0, 5) = "사용수량"
            grid_PartHistory(0, 6) = "Loss수량"
            grid_PartHistory(0, 7) = "투입시간"
            grid_PartHistory(0, 8) = "투입 작업자"
            grid_PartHistory(0, 9) = "종료시간"
            grid_PartHistory(0, 10) = "종료 작업자"
            grid_PartHistory(0, 11) = "확인결과"
            grid_PartHistory(0, 12) = "사유"
            grid_PartHistory(0, 13) = "확인 관리자"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
            .ExtendLastCol = False
            .Cols(4).Format = "#,##0"
            .Cols(5).Format = "#,##0"
            .Cols(6).Format = "#,##0"
            .Cols(7).Format = "yyyy-MM-dd HH:mm:ss"
            .Cols(9).Format = "yyyy-MM-dd HH:mm:ss"
        End With

        'set up styles
        Dim s As CellStyle = grid_PartHistory.Styles(CellStyleEnum.Subtotal0)
        s.BackColor = Color.DarkBlue
        s.ForeColor = Color.White
        s.Font = New Font(grid_PartHistory.Font, FontStyle.Bold)

        s = grid_PartHistory.Styles(CellStyleEnum.Subtotal1)
        s.BackColor = Color.LightBlue
        s.ForeColor = Color.Black
        s = grid_PartHistory.Styles(CellStyleEnum.Subtotal2)
        s.BackColor = Color.DarkRed
        s.ForeColor = Color.White

        'more setup
        grid_PartHistory.AllowDragging = AllowDraggingEnum.None
        grid_PartHistory.AllowEditing = False
        'C1FlexGrid1.Cols(0).WidthDisplay \= 3
        grid_PartHistory.Tree.Column = 1
        grid_PartHistory.SubtotalPosition = SubtotalPositionEnum.AboveData

        With C1FlexGrid1
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
            C1FlexGrid1(0, 0) = "No"
            C1FlexGrid1(0, 1) = "공정명"
            C1FlexGrid1(0, 2) = "프로그램명"
            C1FlexGrid1(0, 3) = "확인일자"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

        With C1FlexGrid2
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
            C1FlexGrid2(0, 0) = "No"
            C1FlexGrid2(0, 1) = "Feeder No."
            C1FlexGrid2(0, 2) = "용도"
            C1FlexGrid2(0, 3) = "결과"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

    End Sub

    Private Sub btn_Search_Click(sender As Object, e As EventArgs) Handles btn_Search.Click

        thread_LoadingFormStart()

        grid_LotList.Rows.Count = 1
        grid_LotList.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_smt_manager(4" &
                                ", '" & Format(Start_Date.Value, "yyyy-MM-dd 00:00:00") & "'" &
                                ", '" & Format(End_Date.Value, "yyyy-MM-dd 23:59:59") & "'" &
                                ", '" & tb_Slip_No.Text & "'" &
                                 ", null" &
                                ")"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insertString As String = grid_LotList.Rows.Count & vbTab &
                sqlDR("lot_no") & vbTab &
                sqlDR("lot_option") & vbTab &
                sqlDR("yj_no") & vbTab &
                sqlDR("work_section") & vbTab &
                Format(sqlDR("end_date"), "yyyy-MM-dd HH:mm:ss")
            grid_LotList.AddItem(insertString)
        Loop
        sqlDR.Close()

        DBClose()

        grid_LotList.AutoSizeCols()
        grid_LotList.Redraw = True

        th_LoadingWindow.Abort()

    End Sub

    Private Sub grid_LotList_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles grid_LotList.MouseDoubleClick

        If e.Button = MouseButtons.Left And grid_LotList.MouseRow > 0 Then
            Dim selRow As Integer = grid_LotList.MouseRow
            thread_LoadingFormStart()
            load_Material_History(grid_LotList(selRow, 1), grid_LotList(selRow, 4))
            load_PGM_Check_History(grid_LotList(selRow, 1), grid_LotList(selRow, 4))
            load_Dipping_Feeder_History(grid_LotList(selRow, 1), grid_LotList(selRow, 4))
            th_LoadingWindow.Abort()
        End If

    End Sub

    Private Sub load_Dipping_Feeder_History(ByVal lot_no As String, ByVal work_setion As String)

        C1FlexGrid2.Rows.Count = 1
        C1FlexGrid2.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_smt_manager(8" &
                                ", null" &
                                ", null" &
                                ", '" & lot_no & "'" &
                                ", '" & work_setion & "'" &
                                ")"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insertString As String = C1FlexGrid2.Rows.Count & vbTab &
                sqlDR("dipping_feeder_no") & vbTab &
                sqlDR("feeder_for_use") & vbTab &
                sqlDR("dipping_feeder_check")
            C1FlexGrid2.AddItem(insertString)
        Loop
        sqlDR.Close()

        DBClose()

        C1FlexGrid2.AutoSizeCols()
        C1FlexGrid2.Redraw = True

    End Sub

    Private Sub load_PGM_Check_History(ByVal lot_no As String, ByVal work_setion As String)

        C1FlexGrid1.Rows.Count = 1
        C1FlexGrid1.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_smt_manager(7" &
                                ", null" &
                                ", null" &
                                ", '" & lot_no & "'" &
                                ", '" & work_setion & "'" &
                                ")"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insertString As String = C1FlexGrid1.Rows.Count & vbTab &
                sqlDR("process_name") & vbTab &
                sqlDR("process_pgm_name") & vbTab &
                Format(sqlDR("check_date"), "yyyy-MM-dd HH:mm:ss")
            C1FlexGrid1.AddItem(insertString)
        Loop
        sqlDR.Close()

        DBClose()

        C1FlexGrid1.AutoSizeCols()
        C1FlexGrid1.Redraw = True

    End Sub

    Private Sub load_Material_History(ByVal lot_no As String, ByVal work_setion As String)

        grid_PartHistory.Rows.Count = 1
        grid_PartHistory.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_smt_manager(5" &
                                ", null" &
                                ", null" &
                                ", '" & lot_no & "'" &
                                ", '" & work_setion & "'" &
                                ")"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim end_date As String = String.Empty
            If IsDBNull(sqlDR("end_date")) Then
                end_date = ""
            Else
                end_date = Format(sqlDR("end_date"), "yyyy-MM-dd HH:mm:ss")
            End If
            Dim insertString As String = grid_PartHistory.Rows.Count & vbTab &
                sqlDR("basic_partno") & vbTab &
                sqlDR("material_part_no") & vbTab &
                sqlDR("material_lot_no") & vbTab &
                Format(sqlDR("basic_stock_qty"), "#,##0") & vbTab &
                Format(sqlDR("material_used_qty"), "#,##0") & vbTab &
                Format(sqlDR("material_loss_qty"), "#,##0") & vbTab &
                Format(sqlDR("start_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                sqlDR("start_worker") & vbTab &
                end_date & vbTab &
                sqlDR("end_worker") & vbTab &
                sqlDR("check_result") & vbTab &
                sqlDR("fault_reason") & vbTab &
                sqlDR("fault_checker")
            grid_PartHistory.AddItem(insertString)
        Loop
        sqlDR.Close()

        DBClose()

        UpdateTotals()

        grid_PartHistory.AutoSizeCols()
        grid_PartHistory.Redraw = True

    End Sub

    Private Sub UpdateTotals()
        'clear existing totals
        grid_PartHistory.Subtotal(AggregateEnum.Clear)

        'calculate subtotals (three levels, totals on every column)
        'Dim c As Integer
        'For c = 3 To C1FlexGrid1.Cols.Count - 1
        'grid_PartHistory.Subtotal(AggregateEnum.Sum, 0, 1, 4, "Grand Total")
        grid_PartHistory.Subtotal(AggregateEnum.Sum, 0, 1, 5, "Grand Total")
        grid_PartHistory.Subtotal(AggregateEnum.Sum, 0, 1, 6, "Grand Total")
        'grid_PartHistory.Subtotal(AggregateEnum.Sum, 1, 3, 4, "Total for {0}")
        grid_PartHistory.Subtotal(AggregateEnum.Sum, 1, 3, 5, "Total for {0}")
        grid_PartHistory.Subtotal(AggregateEnum.Sum, 1, 3, 6, "Total for {0}")
        'C1FlexGrid1.Subtotal(AggregateEnum.Sum, 2, 2, 5, "Total for {0}")
        'Next

        'done, autosize columns to finish
        grid_PartHistory.AutoSizeCols()

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub ToolStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub
End Class
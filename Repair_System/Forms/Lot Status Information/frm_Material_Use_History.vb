Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_Material_Use_History

    Public lot_no As String

    Private Sub frm_Material_Use_History_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        grid_Setting()

        load_PartHistory()


    End Sub

    Private Sub grid_Setting()

        With grid_Lot_List
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 11
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_Lot_List(0, 0) = "No"
            grid_Lot_List(0, 1) = "사용 Lot No."
            grid_Lot_List(0, 2) = "Part No."
            grid_Lot_List(0, 3) = "Lot No."
            grid_Lot_List(0, 4) = "기초재고"
            grid_Lot_List(0, 5) = "사용수량"
            grid_Lot_List(0, 6) = "Loss수량"
            grid_Lot_List(0, 7) = "투입시간"
            grid_Lot_List(0, 8) = "투입 작업자"
            grid_Lot_List(0, 9) = "종료시간"
            grid_Lot_List(0, 10) = "종료 작업자"
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
        End With

        'set up styles
        Dim s As CellStyle = grid_Lot_List.Styles(CellStyleEnum.Subtotal0)
        s.BackColor = Color.DarkBlue
        s.ForeColor = Color.White
        s.Font = New Font(grid_Lot_List.Font, FontStyle.Bold)

        s = grid_Lot_List.Styles(CellStyleEnum.Subtotal1)
        s.BackColor = Color.LightBlue
        s.ForeColor = Color.Black
        s = grid_Lot_List.Styles(CellStyleEnum.Subtotal2)
        s.BackColor = Color.DarkRed
        s.ForeColor = Color.White

        'more setup
        grid_Lot_List.AllowDragging = AllowDraggingEnum.None
        grid_Lot_List.AllowEditing = False
        'C1FlexGrid1.Cols(0).WidthDisplay \= 3
        grid_Lot_List.Tree.Column = 1
        grid_Lot_List.SubtotalPosition = SubtotalPositionEnum.AboveData

    End Sub

    Private Sub load_PartHistory()

        thread_LoadingFormStart()

        grid_Lot_List.Rows.Count = 1
        grid_Lot_List.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_total_material_status(1" &
                                ", null" &
                                ", null" &
                                ", null" &
                                ", null" &
                                ", '" & lot_no & "'" &
                                ")"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim replaceUsedQty As String = Nothing
            Dim replaceEndDate As String = Nothing
            If Not sqlDR("material_used_qty") = "사용중" Then
                replaceUsedQty = Format(CDbl(sqlDR("material_used_qty")), "#,##0")
                replaceEndDate = Format(CDate(sqlDR("end_date")), "yyyy-MM-dd HH:mm:ss")
            Else
                replaceUsedQty = sqlDR("material_used_qty")
                replaceEndDate = sqlDR("end_date")
            End If
            Dim insertString As String = grid_Lot_List.Rows.Count & vbTab &
                sqlDR("lot_no") & vbTab &
                sqlDR("material_part_no") & vbTab &
                sqlDR("material_lot_no") & vbTab &
                Format(sqlDR("basic_stock_qty"), "#,##0") & vbTab &
                replaceUsedQty & vbTab &
                Format(sqlDR("material_loss_qty"), "#,##0") & vbTab &
                Format(sqlDR("start_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                sqlDR("start_worker") & vbTab &
                replaceEndDate & vbTab &
                sqlDR("end_worker")
            grid_Lot_List.AddItem(insertString)
        Loop
        sqlDR.Close()

        DBClose()

        UpdateTotals()

        grid_Lot_List.AutoSizeCols()
        grid_Lot_List.Redraw = True

        th_LoadingWindow.Abort()

    End Sub

    Private Sub UpdateTotals()
        'clear existing totals
        grid_Lot_List.Subtotal(AggregateEnum.Clear)

        'calculate subtotals (three levels, totals on every column)
        'Dim c As Integer
        'For c = 3 To C1FlexGrid1.Cols.Count - 1
        grid_Lot_List.Subtotal(AggregateEnum.Sum, 0, -1, 5, "Grand Total")
        grid_Lot_List.Subtotal(AggregateEnum.Sum, 0, -1, 6, "Grand Total")
        'grid_Lot_List.Subtotal(AggregateEnum.Sum, 0, 1, 5, "Grand Total")
        'grid_Lot_List.Subtotal(AggregateEnum.Sum, 0, 1, 6, "Grand Total")
        'grid_PartHistory.Subtotal(AggregateEnum.Sum, 1, 3, 4, "Total for {0}")
        'grid_Lot_List.Subtotal(AggregateEnum.Sum, 1, -1, 5, "Total for {0}")
        'grid_Lot_List.Subtotal(AggregateEnum.Sum, 1, -1, 6, "Total for {0}")
        'C1FlexGrid1.Subtotal(AggregateEnum.Sum, 2, 2, 5, "Total for {0}")
        'Next

        'done, autosize columns to finish
        grid_Lot_List.AutoSizeCols()

    End Sub

    Private Sub grid_Lot_List_MouseMove(sender As Object, e As MouseEventArgs) Handles grid_Lot_List.MouseMove

        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then

            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub

    Private Sub grid_LotList_LostFocus(sender As Object, e As EventArgs) Handles grid_Lot_List.LostFocus

        frm_Main.lb_Status.Text = String.Empty

    End Sub

    Private Sub frm_Material_Use_History_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        Me.Dispose()

    End Sub

    Private Sub grid_Lot_List_KeyDown(sender As Object, e As KeyEventArgs) Handles grid_Lot_List.KeyDown

        If e.Shift Then
            md_ETC.cellCal(sender, Nothing)
        Else
            frm_Main.lb_Status.Text = String.Empty
        End If

    End Sub
End Class
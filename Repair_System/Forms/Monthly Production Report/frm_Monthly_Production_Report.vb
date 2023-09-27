'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-08-31
'##### 수정일자 : 2023-08-31
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 

'############################################################################################################
'############################################################################################################

Imports System.Globalization
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_Monthly_Production_Report

    Dim form_Msgbox_String As String = "OMS 생산월보"

    Private Sub frm_Monthly_Production_Report_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GRID_SETTING()
        Label1.Visible = False

    End Sub

    Private Sub GRID_SETTING()

        With grid_SlipNoDetail
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 16
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_SlipNoDetail(0, 0) = "No"
            grid_SlipNoDetail(0, 1) = "Product"
            grid_SlipNoDetail(0, 2) = "Fab Line"
            grid_SlipNoDetail(0, 3) = "Lot Type"
            grid_SlipNoDetail(0, 4) = "Return Type"
            grid_SlipNoDetail(0, 5) = Format(DateAdd(DateInterval.Day, -1, DateTimePicker1.Value), "MM월dd일" & vbCrLf & "OMS수량")
            grid_SlipNoDetail(0, 6) = Format(DateTimePicker1.Value, "MM월dd일" & vbCrLf & "우익->유진")
            grid_SlipNoDetail(0, 7) = Format(DateTimePicker1.Value, "MM월dd일" & vbCrLf & "유진->우익")
            grid_SlipNoDetail(0, 8) = Format(DateTimePicker1.Value, "MM월dd일 22시" & vbCrLf & "예상 OMS수량")
            grid_SlipNoDetail(0, 9) = "유진재공수량"
            grid_SlipNoDetail(0, 10) = "OMS수량-재공수량"
            grid_SlipNoDetail(0, 11) = "누적" & vbCrLf & "유진입고수량"
            grid_SlipNoDetail(0, 12) = "누적" & vbCrLf & "유진출고수량"
            grid_SlipNoDetail(0, 13) = Format(DateTimePicker1.Value, "MM월" & vbCrLf & "유진입고수량")
            grid_SlipNoDetail(0, 14) = Format(DateTimePicker1.Value, "MM월" & vbCrLf & "유진출고수량")
            grid_SlipNoDetail(0, 15) = GetWeekOfYear(DateTimePicker1.Value) & "주차" & vbCrLf & "유진출고수량"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .AutoSizeCols()
            .ShowCursor = True
            .SubtotalPosition = SubtotalPositionEnum.AboveData
            For i = 5 To 15
                .Cols(i).Format = "#,##0"
            Next
        End With

        'set up styles
        Dim s As CellStyle = grid_SlipNoDetail.Styles(CellStyleEnum.GrandTotal)
        s.BackColor = Color.DarkBlue
        s.ForeColor = Color.White
        s.Font = New Font(grid_SlipNoDetail.Font, FontStyle.Bold)

        s = grid_SlipNoDetail.Styles(CellStyleEnum.Subtotal0)
        s.BackColor = Color.LightBlue
        s.ForeColor = Color.Black
        s = grid_SlipNoDetail.Styles(CellStyleEnum.Subtotal1)
        s.BackColor = Color.DarkRed
        s.ForeColor = Color.White

    End Sub

    Private Sub btn_Search_Click(sender As Object, e As EventArgs) Handles btn_Search.Click

        thread_LoadingFormStart()

        GRID_SETTING()

        grid_SlipNoDetail.Rows.Count = 1
        grid_SlipNoDetail.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_module_total_information('"
        strSQL += Format(DateTimePicker1.Value, "yyyy-MM-dd") & "')"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insertString As String = grid_SlipNoDetail.Rows.Count & vbTab &
                sqlDR("product") & vbTab &
                sqlDR("fab_line") & vbTab &
                sqlDR("lot_type") & vbTab &
                sqlDR("return_type") & vbTab &
                Format(sqlDR("yesterday_oms"), "#,##0") & vbTab &
                Format(sqlDR("wi_send"), "#,##0") & vbTab &
                Format(sqlDR("uj_send"), "#,##0") & vbTab &
                Format(sqlDR("today_oms"), "#,##0") & vbTab &
                Format(sqlDR("uj_stock_qty"), "#,##0") & vbTab &
                Format(sqlDR("diff_oms_uj"), "#,##0") & vbTab &
                Format(sqlDR("total_uj_in"), "#,##0") & vbTab &
                Format(sqlDR("total_uj_out"), "#,##0") & vbTab &
                Format(sqlDR("month_uj_in"), "#,##0") & vbTab &
                Format(sqlDR("month_uj_out"), "#,##0") & vbTab &
                Format(sqlDR("week_uj_out"), "#,##0")
            grid_SlipNoDetail.AddItem(insertString)
        Loop
        sqlDR.Close()

        DBClose()

        UpdateTotals()

        grid_SlipNoDetail.AutoSizeCols()
        grid_SlipNoDetail.Redraw = True

        Label1.Text = Format(Now, "yyyy-MM-dd HH:mm:ss") & " 조회된 내용입니다."
        Label1.Visible = True

        th_LoadingWindow.Abort()

    End Sub

    Private Sub UpdateTotals()
        'clear existing totals
        grid_SlipNoDetail.Subtotal(AggregateEnum.Clear)

        'calculate subtotals (three levels, totals on every column)
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, -1, 5, "Grand Total")
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, -1, 6, "Grand Total")
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, -1, 7, "Grand Total")
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, -1, 8, "Grand Total")
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, -1, 9, "Grand Total")
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, -1, 10, "Grand Total")
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, -1, 11, "Grand Total")
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, -1, 12, "Grand Total")
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, -1, 13, "Grand Total")
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, -1, 14, "Grand Total")
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, -1, 15, "Grand Total")
        'grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, 1, 6, "Grand Total")
        'grid_PartHistory.Subtotal(AggregateEnum.Sum, 1, 3, 4, "Total for {0}")
        'grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 1, 3, 5, "Total for {0}")
        'grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 1, 3, 6, "Total for {0}")

        'done, autosize columns to finish
        grid_SlipNoDetail.AutoSizeCols()

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub grid_SlipNoDetail_MouseMove(sender As Object, e As MouseEventArgs) Handles grid_SlipNoDetail.MouseMove

        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then

            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub

    Private Sub grid_LotList_LostFocus(sender As Object, e As EventArgs) Handles grid_SlipNoDetail.LostFocus

        frm_Main.lb_Status.Text = String.Empty

    End Sub

    Private Sub grid_SlipNoDetail_KeyDown(sender As Object, e As KeyEventArgs) Handles grid_SlipNoDetail.KeyDown

        If e.Shift Then
            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub
End Class
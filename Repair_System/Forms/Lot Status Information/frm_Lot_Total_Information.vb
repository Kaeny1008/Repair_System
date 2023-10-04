'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-08-10
'##### 수정일자 : 2023-08-10
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 

'############################################################################################################
'############################################################################################################

Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_Lot_Total_Information

    Dim form_Msgbox_String As String = "Lot 정보보기"
    Dim first_run As Integer = 0

    Private Sub frm_Lot_Total_Information_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'SplitContainer1.Panel1Collapsed = True

        grid_Setting()

        'lb_Total.Focus()

    End Sub

    Private Sub grid_Setting()

        With grid_Lot_List
            .AllowEditing = False
            .AllowFiltering = True
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 17
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_Lot_List(0, 0) = "No"
            grid_Lot_List(0, 1) = "YJ No."
            grid_Lot_List(0, 2) = "Slip No."
            grid_Lot_List(0, 3) = "Product"
            grid_Lot_List(0, 4) = "Lot No."
            grid_Lot_List(0, 5) = "Lot Option"
            grid_Lot_List(0, 6) = "Module Qty"
            grid_Lot_List(0, 7) = "Lot 상태"
            grid_Lot_List(0, 8) = "입고일자"
            grid_Lot_List(0, 9) = "수입검사 일자"
            grid_Lot_List(0, 10) = "베이킹 일자"
            grid_Lot_List(0, 11) = "현장출고 일자"
            grid_Lot_List(0, 12) = "PMIC 작업일자"
            grid_Lot_List(0, 13) = "RCD 작업일자"
            grid_Lot_List(0, 14) = "출하 일자"
            grid_Lot_List(0, 15) = "폐기자재 회수여부"
            grid_Lot_List(0, 16) = "비고"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .ExtendLastCol = False
            .AutoSizeCols()
            .ShowCursor = True
            .Tree.Style = TreeStyleFlags.Simple
        End With

    End Sub

    Private Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        thread_LoadingFormStart()

        first_run = 1
        grid_Lot_List.Rows.Count = 1
        grid_Lot_List.Redraw = False

        lb_Total.Text = 0
        lb_Inspect_Ready.Text = 0
        lb_Inspection_Completed.Text = 0
        lb_Baking_Completed.Text = 0
        lb_Moving_Completed.Text = 0
        lb_Completed.Text = 0
        Label8.Text = 0

        DBConnect()

        Dim strSQL As String = "call sp_total_lot_status()"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_Lot_List.Rows.Count & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          sqlDR("slip_no") & vbTab &
                                          sqlDR("product") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("lot_option") & vbTab &
                                          Format(sqlDR("chip_qty"), "#,##0") & vbTab &
                                          sqlDR("lot_status") & vbTab &
                                          sqlDR("confirm_date") & vbTab &
                                          sqlDR("inspect_date") & vbTab &
                                          sqlDR("baking_date") & vbTab &
                                          sqlDR("worksite_moving_date") & vbTab &
                                          sqlDR("pmic_end_date") & vbTab &
                                          sqlDR("rcd_end_date") & vbTab &
                                          sqlDR("ship_date") & vbTab &
                                          sqlDR("reject_material") & vbTab &
                                          sqlDR("note")
            lb_Total.Text = Format(CDbl(lb_Total.Text) + sqlDR("chip_qty"), "#,##0")
            Select Case sqlDR("lot_status")
                Case "Confirm Ready"
                    If CheckBox1.Checked Then grid_Lot_List.AddItem(insert_String)
                Case "Ready"
                    If CheckBox2.Checked Then grid_Lot_List.AddItem(insert_String)
                    lb_Inspect_Ready.Text = Format(CDbl(lb_Inspect_Ready.Text) + sqlDR("chip_qty"), "#,##0")
                Case "Incoming Inspection Completed"
                    If CheckBox3.Checked Then grid_Lot_List.AddItem(insert_String)
                    lb_Inspection_Completed.Text = Format(CDbl(lb_Inspection_Completed.Text) + sqlDR("chip_qty"), "#,##0")
                Case "Baking End"
                    If CheckBox4.Checked Then grid_Lot_List.AddItem(insert_String)
                    lb_Baking_Completed.Text = Format(CDbl(lb_Baking_Completed.Text) + sqlDR("chip_qty"), "#,##0")
                Case "Moving to WorkSite"
                    If CheckBox6.Checked Then grid_Lot_List.AddItem(insert_String)
                    lb_Moving_Completed.Text = Format(CDbl(lb_Moving_Completed.Text) + sqlDR("chip_qty"), "#,##0")
                Case "Completed"
                    If CheckBox7.Checked Then grid_Lot_List.AddItem(insert_String)
                    lb_Completed.Text = Format(CDbl(lb_Completed.Text) + sqlDR("chip_qty"), "#,##0")
                Case Else
                    grid_Lot_List.AddItem(insert_String)
            End Select
            Label2.Text = Format(CDbl(lb_Total.Text) - CDbl(lb_Completed.Text), "#,##0")
            Label8.Text = Format(CDbl(Label8.Text) + 1, "#,##0")
        Loop
        sqlDR.Close()

        DBClose()

        grid_Lot_List.AutoSizeCols()
        grid_Lot_List.Redraw = True

        thread_LoadingFormEnd

    End Sub

    Private Sub grid_Lot_List_KeyDown(sender As Object, e As KeyEventArgs) Handles grid_Lot_List.KeyDown
        '17, 70
        If (e.KeyCode And Not Keys.Modifiers) = Keys.F AndAlso e.Modifiers = Keys.Control Then
            frm_Grid_String_Find.selGrid = Me.grid_Lot_List
            frm_Grid_String_Find.first_row = grid_Lot_List.RowSel
            frm_Grid_String_Find.first_col = grid_Lot_List.ColSel
            If Not frm_Grid_String_Find.Visible Then frm_Grid_String_Find.Show()
            frm_Grid_String_Find.Focus()
        End If

        If e.Shift Then
            md_ETC.cellCal(sender, Nothing)
        Else
            frm_Main.lb_Status.Text = String.Empty
        End If

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

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

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged,
            CheckBox2.CheckedChanged,
            CheckBox3.CheckedChanged,
            CheckBox4.CheckedChanged,
            CheckBox6.CheckedChanged,
            CheckBox7.CheckedChanged

        If first_run = 1 Then BTN_Search_Click(Nothing, Nothing)

    End Sub

    Dim grid_sortFlag As SortFlags
    Dim grid_sortCol As Integer = 0

    Private Sub grid_Lot_List_MouseClick(sender As Object, e As MouseEventArgs) Handles grid_Lot_List.MouseClick

        If e.Button = MouseButtons.Left And grid_Lot_List.MouseRow = 0 Then
            If grid_sortCol = grid_Lot_List.MouseCol Then
                If grid_sortFlag = SortFlags.Ascending Then
                    grid_sortFlag = SortFlags.Descending
                Else
                    grid_sortFlag = SortFlags.Ascending
                End If
            Else
                grid_sortFlag = 0
                grid_sortCol = grid_Lot_List.MouseCol
            End If
            grid_Lot_List.Sort(grid_sortFlag, grid_sortCol)
            For i = 1 To grid_Lot_List.Rows.Count - 1
                grid_Lot_List(i, 0) = i
            Next
        End If

    End Sub
End Class
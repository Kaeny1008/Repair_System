'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-07-24
'##### 수정일자 : 2023-07-24
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 입고검사 대기중인 리스트를 조회( 우선순위 LOT 확인 필요에 의한 폼)

'############################################################################################################
'############################################################################################################

Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_IQC_Standby_List

    Dim form_Msgbox_String As String = "수입검사 대기 리스트"

    Private Sub IN_STAND_BY_LIST_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        grid_Setting()

        Start_Date.Value = Format(Now, "2023-08-01 00:00:00")
        End_Date.Value = Format(Now, "yyyy-MM-dd HH:mm:ss")

    End Sub

    Private Sub grid_Setting()

        With GRID_Lot_List
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 8
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            GRID_Lot_List(0, 0) = "No"
            GRID_Lot_List(0, 1) = "구분"
            GRID_Lot_List(0, 2) = "입고일자"
            GRID_Lot_List(0, 3) = "Slip No."
            GRID_Lot_List(0, 4) = "Product Code" & vbCrLf & " / Part No."
            GRID_Lot_List(0, 5) = "Lot No."
            GRID_Lot_List(0, 6) = "Qty"
            GRID_Lot_List(0, 7) = "비고"
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

    Private Sub BTN_Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Search.Click

        thread_LoadingFormStart()

        GRID_Lot_List.Redraw = False
        GRID_Lot_List.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call sp_iqc_standby_list('" & Format(Start_Date.Value, "yyyy-MM-dd") & "', '" &
                                Format(DateAdd(DateInterval.Day, 1, End_Date.Value), "yyyy-MM-dd") & "', '" &
                                CB_Product_Code.Text & "', '" &
                                0 & "')"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = GRID_Lot_List.Rows.Count & vbTab &
                                          sqlDR("part_division") & vbTab &
                                          Format(sqlDR("first_write_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          sqlDR("slip_no") & vbTab &
                                          sqlDR("product") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          Format(sqlDR("chip_qty"), "#,##0") & vbTab &
                                          sqlDR("note")
            GRID_Lot_List.AddItem(insert_String)
        Loop
        sqlDR.Close()

        strSQL = "call sp_iqc_standby_list('" & Format(Start_Date.Value, "yyyy-MM-dd") & "', '" &
                                Format(End_Date.Value, "yyyy-MM-dd") & "', '" &
                                CB_Product_Code.Text & "', '" &
                                1 & "')"

        sqlCmd = New MySqlCommand(strSQL, dbConnection1)
        sqlDR = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = GRID_Lot_List.Rows.Count & vbTab &
                                          sqlDR("part_division") & vbTab &
                                          Format(sqlDR("first_write_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          sqlDR("slip_no") & vbTab &
                                          sqlDR("part_no") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          Format(sqlDR("material_qty"), "#,##0") & vbTab &
                                          sqlDR("material_note")
            GRID_Lot_List.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        Dim s As CellStyle = GRID_Lot_List.Styles(CellStyleEnum.Subtotal0)
        s.BackColor = Color.AliceBlue
        s.ForeColor = Color.Black
        s.Format = "#,##0"
        GRID_Lot_List.SubtotalPosition = SubtotalPositionEnum.BelowData
        GRID_Lot_List.Subtotal(AggregateEnum.Sum, 0, 1, 6, "구분합계")

        GRID_Lot_List.AutoSizeCols()
        GRID_Lot_List.Redraw = True

        th_LoadingWindow.Abort()

    End Sub

    Private Sub Form_CLose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub GRID_Lot_List_MouseMove(sender As Object, e As MouseEventArgs) Handles GRID_Lot_List.MouseMove

        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then

            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub

    Private Sub grid_LotList_LostFocus(sender As Object, e As EventArgs) Handles GRID_Lot_List.LostFocus

        frm_Main.lb_Status.Text = String.Empty

    End Sub
End Class
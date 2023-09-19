'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-08-10
'##### 수정일자 : 2023-08-10
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 

'############################################################################################################
'############################################################################################################

Imports System.Threading
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_Material_Total_Information

    Dim form_Msgbox_String As String = "자재 종합정보"

    Private Sub frm_Material_Total_Information_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        SplitContainer1.Panel1Collapsed = False
        grid_Setting()

    End Sub

    Private Sub grid_Setting()

        With grid_Lot_List
            .AllowEditing = False
            .AllowFiltering = True
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 8
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_Lot_List(0, 0) = "No"
            grid_Lot_List(0, 1) = "Part No."
            grid_Lot_List(0, 2) = "입고수량"
            grid_Lot_List(0, 3) = "사용수량"
            grid_Lot_List(0, 4) = "Loss수량"
            grid_Lot_List(0, 5) = "재고수량"
            grid_Lot_List(0, 6) = "필요수량"
            grid_Lot_List(0, 7) = "차이"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .ExtendLastCol = False
            .AutoSizeCols()
            .ShowCursor = True
            .Tree.Style = TreeStyleFlags.Simple
        End With

    End Sub

    Private Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        thread_LoadingFormStart()

        grid_Lot_List.Rows.Count = 1
        grid_Lot_List.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_parts_required_quantity()"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_Lot_List.Rows.Count & vbTab &
                                          sqlDR("part_no") & vbTab &
                                          Format(sqlDR("total_qty"), "#,##0") & vbTab &
                                          Format(sqlDR("used_qty"), "#,##0") & vbTab &
                                          Format(sqlDR("loss_qty"), "#,##0") & vbTab &
                                          Format(sqlDR("stock_qty"), "#,##0") & vbTab &
                                          Format(sqlDR("required_qty"), "#,##0") & vbTab &
                                          Format(sqlDR("diff_qty"), "#,##0")
            grid_Lot_List.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        grid_Lot_List.AutoSizeCols()
        grid_Lot_List.Redraw = True

        th_LoadingWindow.Abort()

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

    End Sub

    Private Sub grid_Lot_List_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles grid_Lot_List.MouseDoubleClick

        'If e.Button = MouseButtons.Left And grid_Lot_List.MouseRow > 0 Then
        '    Dim selRow As Integer = grid_Lot_List.MouseRow
        '    frm_Material_Use_History.lot_no = grid_Lot_List(selRow, 4)
        '    If Not frm_Material_Use_History.Visible Then frm_Material_Use_History.Show()
        '    frm_Material_Use_History.Focus()
        'End If

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
End Class
'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-09-20
'##### 수정일자 : 2023-09-20
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 

'############################################################################################################
'############################################################################################################

Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_Lot_Residence_Time

    Dim form_Msgbox_String As String = "Lot 정체일(현장출고 후)"

    Private Sub frm_Lot_Residence_Time_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
            .Cols.Count = 9
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_Lot_List(0, 0) = "No"
            grid_Lot_List(0, 1) = "Product"
            grid_Lot_List(0, 2) = "Lot No."
            grid_Lot_List(0, 3) = "YJ No."
            grid_Lot_List(0, 4) = "모듈수량"
            grid_Lot_List(0, 5) = "Lot Option"
            grid_Lot_List(0, 6) = "Lot 상태"
            grid_Lot_List(0, 7) = "현장 출고일자"
            grid_Lot_List(0, 8) = "정체일"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            '.ExtendLastCol = False
            .AutoSizeCols()
            .ShowCursor = True
            .Tree.Style = TreeStyleFlags.Simple
        End With

    End Sub

    Private Sub btn_Search_Click(sender As Object, e As EventArgs) Handles btn_Search.Click

        thread_LoadingFormStart()

        grid_Lot_List.Rows.Count = 1
        grid_Lot_List.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_lot_residence_time(" &
                               "'" & tb_Lot_No.Text & "'" &
                               ");"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_Lot_List.Rows.Count & vbTab &
                                          sqlDR("product") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          Format(sqlDR("chip_qty"), "#,##0") & vbTab &
                                          sqlDR("lot_option") & vbTab &
                                          sqlDR("lot_status") & vbTab &
                                          Format(sqlDR("moving_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          Format(sqlDR("residence_time"), "#,##0")
            grid_Lot_List.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        grid_Lot_List.AutoSizeCols()
        grid_Lot_List.Redraw = True

        th_LoadingWindow.Abort()

    End Sub
End Class
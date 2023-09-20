'############################################################################################################
'############################################################################################################

'##### 폼번호   : 99
'##### 작성일자 : 2019-08-09
'##### 수정일자 : 2019-08-09
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 

'############################################################################################################
'############################################################################################################

Imports System.Threading
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_OMS_WIP_DATA

    Dim form_Msgbox_String As String = "OMS 재공데이터 생성"

    Private Sub OMS_WIP_DATA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        GRID_Setting()

        'Table_DT_Load()

    End Sub

    Private Sub GRID_Setting()

        With GRID_Data_List
            .AllowEditing = True
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 11
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            GRID_Data_List(0, 0) = "No"
            GRID_Data_List(0, 1) = "Field ID"
            GRID_Data_List(0, 2) = "Null"
            GRID_Data_List(0, 3) = "Field Name"
            GRID_Data_List(0, 4) = "Type"
            GRID_Data_List(0, 5) = "Data Length"
            GRID_Data_List(0, 6) = "Accumulation"
            GRID_Data_List(0, 7) = "Description"
            GRID_Data_List(0, 8) = "Data Code"
            GRID_Data_List(0, 9) = "Input Data"
            GRID_Data_List(0, 10) = "Text Length"
            .Cols(8).Visible = False
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 3).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .Cols(.Cols.Count - 2).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

        With GRID_Lot_List
            .AllowEditing = True
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 34
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            GRID_Lot_List(0, 0) = "No"
            GRID_Lot_List(0, 1) = "구분"
            GRID_Lot_List(0, 2) = "Start Tag"
            GRID_Lot_List(0, 3) = "OEM Code"
            GRID_Lot_List(0, 4) = "Product Code"
            GRID_Lot_List(0, 5) = "Run ID"
            GRID_Lot_List(0, 6) = "Lot No."
            GRID_Lot_List(0, 7) = "Lot Type"
            GRID_Lot_List(0, 8) = "Return Type"
            GRID_Lot_List(0, 9) = "Process ID"
            GRID_Lot_List(0, 10) = "Step ID"
            GRID_Lot_List(0, 11) = "Step SEQ NO"
            GRID_Lot_List(0, 12) = "Step Desc"
            GRID_Lot_List(0, 13) = "Step In Datetime"
            GRID_Lot_List(0, 14) = "Area Flag"
            GRID_Lot_List(0, 15) = "Area ID"
            GRID_Lot_List(0, 16) = "Chip Qty"
            GRID_Lot_List(0, 17) = "Wafer Qty"
            GRID_Lot_List(0, 18) = "Item Flag"
            GRID_Lot_List(0, 19) = "Hold Flag"
            GRID_Lot_List(0, 20) = "Hold_Code"
            GRID_Lot_List(0, 21) = "Hold DateTime"
            GRID_Lot_List(0, 22) = "NCF Code"
            GRID_Lot_List(0, 23) = "NCA Code"
            GRID_Lot_List(0, 24) = "NCT Code"
            GRID_Lot_List(0, 25) = "NCQ Code"
            GRID_Lot_List(0, 26) = "Other"
            GRID_Lot_List(0, 27) = "Loss Qty"
            GRID_Lot_List(0, 28) = "Bonus Qty"
            GRID_Lot_List(0, 29) = "Fab Line"
            GRID_Lot_List(0, 30) = "Wafer ID"
            GRID_Lot_List(0, 31) = "Create DateTime"
            GRID_Lot_List(0, 32) = "Cutoff YMD"
            GRID_Lot_List(0, 33) = "Inkless"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 2).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

    End Sub

    Private Sub Table_DT_Load()

        thread_LoadingFormStart()

        GRID_Data_List.Redraw = False
        GRID_Data_List.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "select data_sort, field_id, field_null, field_name, field_type, data_length, accumulation, field_desc, data_code"
        strSQL += " from setting_oms_file_data"
        strSQL += " where file_type = '" & "CMWIPINF" & "'"
        strSQL += " order by data_sort"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = GRID_Data_List.Rows.Count & vbTab &
                                          sqlDR("field_id") & vbTab &
                                          sqlDR("field_null") & vbTab &
                                          sqlDR("field_name") & vbTab &
                                          sqlDR("field_type") & vbTab &
                                          sqlDR("data_length") & vbTab &
                                          sqlDR("accumulation") & vbTab &
                                          sqlDR("field_desc") & vbTab &
                                          sqlDR("data_code")
            GRID_Data_List.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        GRID_Data_List.AutoSizeCols()
        GRID_Data_List.Redraw = True

        th_LoadingWindow.Abort()

    End Sub

    Private Sub BTN_Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Search.Click

        Table_DT_Load()

        Lot_List_Load()

        GRID_Lot_List.TopRow = 0

    End Sub

    Private Sub Lot_List_Load()

        thread_LoadingFormStart()

        GRID_Lot_List.Redraw = False
        GRID_Lot_List.Rows.Count = 1

        DBConnect()

        Dim total_Module As Integer = 0
        Dim total_Lot As Integer = 0

        Dim strSQL As String = "call sp_oms_wip_data"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            If IsDBNull(sqlDR("CONFIRM_DATE")) Then
                th_LoadingWindow.Abort()
                Thread.Sleep(100)
                'MsgBox("입고 확정되지 않은 Lot이 존재합니다.", MsgBoxStyle.Information, form_Msgbox_String)
                MessageBox.Show(Me,
                                "입고 확정되지 않은 Lot이 존재합니다.",
                                form_Msgbox_String,
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
                GRID_Lot_List.Rows.Count = 1
                GRID_Lot_List.Redraw = True
                Exit Sub
            End If

            Dim process_id As String = TB_M_Process_ID.Text
            Dim step_id As String = TB_M_Step_ID.Text
            Dim step_seq_no As String = TB_M_Step_SEQ.Text
            Dim step_desc As String = TB_M_Step_Desc.Text
            Dim area_flag As String = TB_M_Area_Flag.Text
            Dim area_id As String = TB_M_Area_ID.Text


            If Not sqlDR("part_division") = "Module" Then
                process_id = TB_C_Process_ID.Text
                step_id = TB_C_Step_ID.Text
                step_seq_no = TB_C_Step_SEQ.Text
                step_desc = TB_C_Step_Desc.Text
                area_flag = TB_C_Area_Flag.Text
                area_id = TB_C_Area_ID.Text
            End If

            Dim write_date As String = Format(Now, "yyyyMMddHHmmss")
            Dim write_day As String = Format(Now, "yyyyMMdd")
            cttm = Format(Now, "yyyyMMddHH")

            Dim insert_String As String = GRID_Lot_List.Rows.Count & vbTab &
                                          sqlDR("part_division") & vbTab &
                                          TB_Start_Tag.Text & vbTab &
                                          TB_Company_Code.Text & vbTab &
                                          sqlDR("product") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("lot_type") & vbTab &
                                          sqlDR("return_type") & vbTab &
                                          process_id & vbTab &
                                          step_id & vbTab &
                                          step_seq_no & vbTab &
                                          step_desc & vbTab &
                                          Format(sqlDR("confirm_date"), "yyyyMMddHHmmss") & vbTab &
                                          area_flag & vbTab &
                                          area_id & vbTab &
                                          sqlDR("chip_qty") & vbTab &
                                          sqlDR("wafer_qty") & vbTab &
                                          TB_Item_Flag.Text & vbTab &
                                          TB_Hold_Flag.Text & vbTab &
                                          TB_Hold_Code.Text & vbTab &
                                          TB_Hold_Datetime.Text & vbTab &
                                          TB_NCF_Code.Text & vbTab &
                                          TB_NCA_Code.Text & vbTab &
                                          TB_NCT_Code.Text & vbTab &
                                          TB_NCQ_Code.Text & vbTab &
                                          TB_Other.Text & vbTab &
                                          TB_Loss_QTY.Text & vbTab &
                                          TB_Bonus_Qty.Text & vbTab &
                                          sqlDR("fab_line") & vbTab &
                                          TB_Wafer_ID.Text & vbTab &
                                          write_date & vbTab &
                                          write_day & vbTab &
                                          TB_Inkless.Text
            GRID_Lot_List.AddItem(insert_String)
            total_Module += CDbl(sqlDR("chip_qty"))
            total_Lot += 1
        Loop
        sqlDR.Close()

        DBClose()

        GRID_Lot_List.AutoSizeCols()
        GRID_Lot_List.Redraw = True

        Label22.Text = "총 모듈 수량 : " & Format(total_Module, "#,##0") & " EA"
        Label25.Text = "총 Lot 수량 : " & Format(total_Lot, "#,##0") & " Lot"

        th_LoadingWindow.Abort()

    End Sub

    Private Sub BTN_Table_Show_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Table_Show.Click

        If SplitContainer1.Panel1Collapsed = True Then
            SplitContainer1.Panel1Collapsed = False
            BTN_Table_Show.Text = "Table 닫기"
        Else
            SplitContainer1.Panel1Collapsed = True
            BTN_Table_Show.Text = "Table 보기"
        End If

    End Sub

    Dim cttm As String

    Private Sub BTN_Data_Output_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Data_Output.Click,
                                                                                                          BTN_Data_Output2.Click

        If GRID_Lot_List.Rows.Count = 1 Then
            MsgBox("재공 데이터를 먼저 조회하십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        For i = 1 To GRID_Lot_List.Rows.Count - 1
            If GRID_Lot_List(i, 5) Like "*-*" Or GRID_Lot_List(i, 6) Like "*-*" Then
                MsgBox("재입고분이 있습니다. 확인하여 수정하여 주십시오." & vbCrLf & "행 :" & i,
                       MsgBoxStyle.Information,
                       form_Msgbox_String)
                Exit Sub
            End If
        Next

        If MsgBox("재공 데이터파일을 생성하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Writing...")

        Dim saveFile As New SaveFileDialog()
        'CMWIPINF_XXYYYYMMDDHH.ff
        saveFile.Filter = "ff File|*.ff"
        saveFile.FileName = "CMWIPINF_" & TB_Company_Code.Text & cttm & ".ff"
        saveFile.Title = "재공 데이터파일 생성"
        saveFile.InitialDirectory = "D:\1. Workplace\5. 고객사 자료\18. 삼성전자(Repair)\##1. OMS Upload\1. 재공\1. 대기"
        saveFile.ShowDialog()

        If saveFile.FileName = String.Empty Then Exit Sub

        'System.IO.File.Copy(Application.StartupPath & "\TEMP_FILE\WIP_INFO.ff", saveFile.FileName, True)

        'Dim file As System.IO.StreamWriter
        'file = My.Computer.FileSystem.OpenTextFileWriter(saveFile.FileName, True)

        Dim TF As Object
        Dim TFT As Object
        TF = CreateObject("scripting.filesystemobject")
        TFT = TF.Createtextfile(saveFile.FileName)

        For i = 1 To GRID_Lot_List.Rows.Count - 1
            '5 Data Length
            Dim write_line As String = String.Empty
            For j = 2 To GRID_Lot_List.Cols.Count - 1
                Dim textLengh As Integer = GRID_Data_List(j - 1, 5)
                Dim oldData As String = GRID_Lot_List(i, j).ToString
                Dim oldTextlengh As Integer = Len(oldData)
                Dim diffLengh As Integer = textLengh - oldTextlengh
                Dim newText As String = String.Empty

                'newText = oldData

                If Not diffLengh = 0 Then
                    For ii = 1 To diffLengh
                        newText += " "
                    Next
                    oldData += newText
                Else
                    oldData = oldData
                End If
                GRID_Data_List(j - 1, 9) = oldData
                GRID_Data_List(j - 1, 10) = Len(oldData)
                write_line += oldData
            Next

            'file.WriteLine(write_line)
            TFT.WriteLine(write_line)
        Next

        'file.Close()
        'file = Nothing
        TFT.close()
        TF = Nothing
        TFT = Nothing

        cttm = String.Empty

        th_LoadingWindow.Abort()
        Thread.Sleep(100)
        MsgBox("재공 데이터 생성완료.", MsgBoxStyle.Information, form_Msgbox_String)
        Label22.Text = "총 모듈 수량 : "
        Label25.Text = "총 Lot 수량 : "
        GRID_Lot_List.Rows.Count = 1

    End Sub

    Private Sub Form_CLose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub GRID_Lot_List_MouseMove(sender As Object, e As MouseEventArgs) Handles GRID_Lot_List.MouseMove

        Dim tip As String = String.Empty
        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then

            Dim total_Qty As Double = 0
            For i = sender.Selection.TopRow To sender.Selection.BottomRow
                For j = sender.Selection.LeftCol To sender.Selection.RightCol
                    If IsNumeric(sender(i, j)) Then
                        total_Qty += sender(i, j)
                    End If
                Next
            Next

            If Not total_Qty = 0 Then _
                ToolTip1.SetToolTip(sender, "합계 : " & Format(total_Qty, "#,##0"))
        End If

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click

        For i = 1 To GRID_Lot_List.Rows.Count - 1
            If GRID_Lot_List(i, 5) Like "*-*" Or GRID_Lot_List(i, 6) Like "*-*" Then
                GRID_Lot_List(i, 5) = Replace(GRID_Lot_List(i, 5), "-1", String.Empty)
                GRID_Lot_List(i, 6) = Replace(GRID_Lot_List(i, 6), "-1", String.Empty)
            End If
        Next

    End Sub

    Private Sub GRID_Lot_List_MouseClick(sender As Object, e As MouseEventArgs) Handles GRID_Lot_List.MouseClick

        If e.Button = MouseButtons.Right And GRID_Lot_List.MouseRow > 0 Then
            Dim selRow As Integer = GRID_Lot_List.MouseRow
            GRID_Lot_List.Row = selRow
            cms_Menu.Show(GRID_Lot_List, New Point(e.X, e.Y))
        End If

    End Sub

    Private Sub btn_RowDelete_Click(sender As Object, e As EventArgs) Handles btn_RowDelete.Click

        GRID_Lot_List.Redraw = False
        GRID_Lot_List.RemoveItem(GRID_Lot_List.Row)

        Dim totalChip As Integer = 0

        For i = 1 To GRID_Lot_List.Rows.Count - 1
            GRID_Lot_List(i, 0) = i
            totalChip += CInt(GRID_Lot_List(i, 16))
        Next

        GRID_Lot_List.Redraw = True

        MsgBox("조회를 완료 하였습니다." & vbCrLf &
            "총 Module 수량 : " & Format(totalChip, "#,##0") & " EA" & vbCrLf &
            "데이터를 생성하여 주십시오.",
            MsgBoxStyle.Information,
            form_Msgbox_String)

    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click

    End Sub
End Class
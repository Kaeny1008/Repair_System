'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-08-07
'##### 수정일자 : 2023-08-07
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 

'############################################################################################################
'############################################################################################################

Imports System.Threading
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_OMS_Ship_Data

    Dim form_Msgbox_String As String = "OMS 반제품 출고 데이터 생성"

    Private Sub frm_OMS_Ship_Data_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GRID_Setting()

        SplitContainer1.Panel1Collapsed = True

        TextBox4.SelectAll()
        TextBox4.Focus()

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
            .Cols.Count = 31
            .Cols.Fixed = 1
            .Rows.Count = 2
            .Rows.Fixed = 1
            GRID_Lot_List(0, 0) = "No"
            GRID_Lot_List(0, 1) = "Start Tag"
            GRID_Lot_List(0, 2) = "Company Code"
            GRID_Lot_List(0, 3) = "Issue Date"
            GRID_Lot_List(0, 4) = "Slip Number"
            GRID_Lot_List(0, 5) = "Product Code"
            GRID_Lot_List(0, 6) = "Sales Code"
            GRID_Lot_List(0, 7) = "Run ID"
            GRID_Lot_List(0, 8) = "Lot ID"
            GRID_Lot_List(0, 9) = "Lot Type"
            GRID_Lot_List(0, 10) = "Return Type"
            GRID_Lot_List(0, 11) = "Wafer Qty"
            GRID_Lot_List(0, 12) = "Chip Qty"
            GRID_Lot_List(0, 13) = "Loss Qty"
            GRID_Lot_List(0, 14) = "Work Week"
            GRID_Lot_List(0, 15) = "Process ID"
            GRID_Lot_List(0, 16) = "Customer Name"
            GRID_Lot_List(0, 17) = "PCB Vendor"
            GRID_Lot_List(0, 18) = "Check Option"
            GRID_Lot_List(0, 19) = "From Area"
            GRID_Lot_List(0, 20) = "To Area"
            GRID_Lot_List(0, 21) = "Match Field"
            GRID_Lot_List(0, 22) = "File Creation DateTime"
            GRID_Lot_List(0, 23) = "Transfer Permit Number"
            GRID_Lot_List(0, 24) = "Assembly Site"
            GRID_Lot_List(0, 25) = "To Site"
            GRID_Lot_List(0, 26) = "Fab Line"
            GRID_Lot_List(0, 27) = "Area Flag"
            GRID_Lot_List(0, 28) = "Module Stock In Location ID"
            GRID_Lot_List(0, 29) = "Mold DateTime"
            GRID_Lot_List(0, 30) = "Inkless"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 2).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

    End Sub

    Private Sub BTN_Table_Show_Click(sender As Object, e As EventArgs) Handles btn_Table_Show.Click

        If SplitContainer1.Panel1Collapsed = True Then
            SplitContainer1.Panel1Collapsed = False
        Else
            SplitContainer1.Panel1Collapsed = True
        End If

    End Sub

    Private Sub Table_DT_Load()

        GRID_Data_List.Redraw = False
        GRID_Data_List.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "select data_sort, field_id, field_null, field_name, field_type, data_length, accumulation, field_desc, data_code"
        strSQL += " from setting_oms_file_data"
        strSQL += " where file_type = '" & "CMHALFIN" & "'"
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

    End Sub

    Dim cttm As String = String.Empty

    Private Sub btn_Search_Click(sender As Object, e As EventArgs) Handles btn_Search.Click

        If Trim(TextBox4.Text) = String.Empty Then
            MsgBox("Slip No.를 입력하여 주십시오.",
                   MsgBoxStyle.Information,
                   form_Msgbox_String)
            Exit Sub
        End If

        thread_LoadingFormStart()

        Table_DT_Load()

        GRID_Lot_List.Redraw = False
        GRID_Lot_List.Rows.Count = 1

        Dim total_Lot As Integer = 0
        Dim total_Module As Integer = 0

        DBConnect()

        Dim strSQL As String = "call sp_oms_ship_data('" & TextBox4.Text & "')"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim write_date As String = Format(Now, "yyyyMMddHHmmss")
            Dim write_day As String = Format(Now, "yyyyMMdd")
            cttm = Format(Now, "yyyyMMddHH")

            Dim insert_String As String = GRID_Lot_List.Rows.Count & vbTab &
                                          tb_Start_Tag.Text & vbTab &
                                          tb_Company_Code.Text & vbTab &
                                          sqlDR("issue_date") & vbTab &
                                          sqlDR("slip_no") & vbTab &
                                          sqlDR("product") & vbTab &
                                          sqlDR("sales_code") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("lot_type") & vbTab &
                                          sqlDR("return_type") & vbTab &
                                          sqlDR("wafer_qty") & vbTab &
                                          sqlDR("chip_qty") & vbTab &
                                          CDbl(tb_Loss_Qty.Text) & vbTab &
                                          tb_Work_Week.Text & vbTab &
                                          tb_Process_ID.Text & vbTab &
                                          tb_Customer_Name.Text & vbTab &
                                          tb_PCB_Name.Text & vbTab &
                                          tb_Check_Option.Text & vbTab &
                                          tb_From_Area.Text & vbTab &
                                          tb_To_Area.Text & vbTab &
                                          tb_Match_Field.Text & vbTab &
                                          write_date & vbTab &
                                          tb_TPN_No.Text & vbTab &
                                          tb_Assembly_Site.Text & vbTab &
                                          tb_To_Site.Text & vbTab &
                                          sqlDR("fab_line") & vbTab &
                                          tb_Area_Flag.Text & vbTab &
                                          tb_MDL_LOC_ID.Text & vbTab &
                                          tb_Mold_DateTime.Text & vbTab &
                                          tb_Inkless.Text
            GRID_Lot_List.AddItem(insert_String)
            total_Lot += 1
            total_Module += CInt(sqlDR("chip_qty"))
        Loop
        sqlDR.Close()

        DBClose()

        GRID_Lot_List.AutoSizeCols()
        GRID_Lot_List.Redraw = True

        GRID_Lot_List.TopRow = 0

        Label9.Text = "총 출고 Lot : " & Format(total_Lot, "#,##0") & " EA"
        Label10.Text = "총 출고 Module : " & Format(total_Module, "#,##0") & " EA"

        th_LoadingWindow.Abort()

    End Sub

    Private Sub btn_Data_Output_Click(sender As Object, e As EventArgs) Handles btn_Data_Output.Click

        If GRID_Lot_List.Rows.Count = 1 Then
            MsgBox("재공 데이터를 먼저 조회하십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        For i = 1 To GRID_Lot_List.Rows.Count - 1
            If GRID_Lot_List(i, 8) Like "*-*" Then
                MsgBox("재입고분이 있습니다. 확인하여 수정하여 주십시오." & vbCrLf & "행 :" & i,
                       MsgBoxStyle.Information,
                       form_Msgbox_String)
                Exit Sub
            End If
        Next

        If MsgBox("반제품 출고 데이터파일을 생성하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Writing...")

        Dim saveFile As New SaveFileDialog()
        saveFile.Filter = "ff File|*.ff"
        saveFile.FileName = "CMHALFIN_" & tb_Company_Code.Text & cttm & ".ff"
        saveFile.Title = "재공 데이터파일 생성"
        saveFile.InitialDirectory = "D:\1. Workplace\5. 고객사 자료\18. 삼성전자(Repair)\##1. OMS Upload\2. 출하\1. 대기"
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
            For j = 1 To GRID_Lot_List.Cols.Count - 1
                Dim textLengh As Integer = GRID_Data_List(j, 5)
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
                GRID_Data_List(j, 9) = oldData
                GRID_Data_List(j, 10) = Len(oldData)
                write_line += oldData
            Next

            'file.WriteLine(write_line)
            'Console.WriteLine(write_line)
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
        MsgBox("반제품 출고 데이터 생성완료.", MsgBoxStyle.Information, form_Msgbox_String)
        Label9.Text = "총 출고 Lot : "
        Label10.Text = "총 출고 Module : "
        GRID_Lot_List.Rows.Count = 1

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

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
            If GRID_Lot_List(i, 8) Like "*-*" Then
                GRID_Lot_List(i, 8) = Replace(GRID_Lot_List(i, 8), "-1", String.Empty)
            End If
        Next

    End Sub
End Class
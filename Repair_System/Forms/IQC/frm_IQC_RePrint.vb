'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-07-13
'##### 수정일자 : 2023-07-13
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 입고된 Lot의 정보를 확인하고 YJ No.를 매칭시켜 라벨을 발행하는 폼

'############################################################################################################
'############################################################################################################

Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Text
Imports System.Threading

Public Class frm_IQC_RePrint

    Dim form_Msgbox_String As String = "수입검사 및 라벨발행"
    Private WithEvents ComPort As New Ports.SerialPort
    Dim c_product_code As String
    Dim c_qty As Integer
    Dim c_fab_line As String
    Dim c_ww As String
    Dim c_option As String
    Dim part_top_marking As String
    Dim nowProduct As String
    Dim nowChipQty As Integer
    Dim first_run As Boolean = True

    Private Sub IN_N_PRINT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        SplitContainer1.Panel1Collapsed = True

        grid_Setting()
        Setting_LOAD()

        tb_lotNo.SelectAll()
        tb_lotNo.Focus()

        history_Load()

    End Sub

    Private Sub grid_Setting()

        With grid_History
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 10
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_History(0, 0) = "No"
            grid_History(0, 1) = "Lot No."
            grid_History(0, 2) = "YJ No."
            grid_History(0, 3) = "공정명"
            grid_History(0, 4) = "공정 번호"
            grid_History(0, 5) = "발행자"
            grid_History(0, 6) = "발행일"
            grid_History(0, 7) = "발행방법"
            grid_History(0, 8) = "발행사유"
            grid_History(0, 9) = "요청자"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

    End Sub

    Private Sub history_Load()

        thread_LoadingFormStart()

        grid_History.Redraw = False
        grid_History.Rows.Count = 1

        DBConnect()

        Try
            Dim strSQL As String = "call sp_label_reprint(null, null, 2)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                Dim insert_string As String = grid_History.Rows.Count & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          sqlDR("process_char") & vbTab &
                                          sqlDR("process_no") & vbTab &
                                          sqlDR("printer") & vbTab &
                                          Format(sqlDR("print_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          sqlDR("print_method") & vbTab &
                                          sqlDR("print_reason") & vbTab &
                                          sqlDR("print_requester")
                grid_History.AddItem(insert_string)
            Loop
            sqlDR.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
        End Try
        DBClose()

        grid_History.AutoSizeCols()
        grid_History.Redraw = True

        thread_LoadingFormEnd

    End Sub

    Private Sub cb_Work_Side_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_Process.SelectedIndexChanged

        Select Case cb_Process.Text
            Case "제전박스(수입검사)"
                tb_Process_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "IQC", "I")
            Case "Baking JIG"
                tb_Process_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "Baking", "B")
            Case "Tray Box(Bake)"
                tb_Process_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "Bake", "A")
            Case "Carrier JIG"
                tb_Process_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "Carrier", "C")
            Case "Tray Box(외관불량)"
                tb_Process_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "X-Ray", "X")
            Case "Tray Box(기능불량)"
                tb_Process_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "Box", "O")
            Case "폐기자재"
                tb_Process_Char.Text = "R"
        End Select

        DBConnect()

        Try
            Dim strSQL As String = "call sp_label_reprint('" & tb_lotNo.Text & "', '" & tb_Process_Char.Text & "', 1)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                tb_total_print_qty.Text = sqlDR("total_print_qty")
            Loop
            sqlDR.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
        End Try
        DBClose()

    End Sub

    Private Sub Setting_LOAD()

        cb_Cable.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "COM/LPT", "COM")
        cb_PrinterList.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Printer Name", "")
        TB_Port.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Port", 1)
        TB_LEFT_Loc.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Width", 0)
        TB_TOP_Loc.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Height", 0)
        TB_BaudRate.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "BaudRate", 9600)
        TB_DataBits.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "DataBits", 8)
        TB_Parity.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Parity", 0)
        TB_StopBits.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "StopBits", 1)

    End Sub

    Private Sub tb_lotNo_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_lotNo.KeyDown

        If e.KeyCode = 13 And Not tb_lotNo.Text = String.Empty Then
            tb_lotNo.Text = tb_lotNo.Text.ToUpper
            tb_Print_Lot_No.Text = String.Empty
            tb_Print_YJ_No.Text = String.Empty
            tb_before_Bucket.Text = String.Empty
            tb_After_Bucket.Text = String.Empty
            tb_Before_PMIC.Text = String.Empty
            tb_Before_PMIC_Marking.Text = String.Empty
            tb_After_PMIC.Text = String.Empty
            tb_After_PMIC_Marking.Text = String.Empty
            tb_Before_RCD.Text = String.Empty
            tb_Before_RCD_Marking.Text = String.Empty
            tb_After_RCD.Text = String.Empty
            tb_After_RCD_Marking.Text = String.Empty

            DBConnect()

            Try
                Dim strSQL As String = "call sp_label_reprint('" & tb_lotNo.Text & "', null, 0)"

                Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
                Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

                Do While sqlDR.Read
                    tb_Print_Lot_No.Text = sqlDR("lot_no")
                    tb_Print_YJ_No.Text = sqlDR("yj_no")
                    tb_before_Bucket.Text = sqlDR("before_bucket")
                    tb_After_Bucket.Text = sqlDR("after_bucket")
                    tb_Before_PMIC.Text = sqlDR("before_pmic_partno")
                    tb_Before_PMIC_Marking.Text = sqlDR("before_pmic_top_marking")
                    tb_After_PMIC.Text = sqlDR("after_pmic_partno")
                    tb_After_PMIC_Marking.Text = sqlDR("after_pmic_top_marking")
                    tb_Before_RCD.Text = sqlDR("before_rcd_partno")
                    tb_Before_RCD_Marking.Text = sqlDR("before_rcd_top_marking")
                    tb_After_RCD.Text = sqlDR("after_rcd_partno")
                    tb_After_RCD_Marking.Text = sqlDR("after_rcd_top_marking")
                    tb_Lot_Option.Text = sqlDR("lot_option")
                    nowProduct = sqlDR("product")
                    nowChipQty = sqlDR("chip_qty")
                Loop
                sqlDR.Close()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            End Try
            DBClose()

            If tb_Print_YJ_No.Text = String.Empty Then
                MsgBox("해당 Lot 정보를 찾을 수 없습니다.", MsgBoxStyle.Information, form_Msgbox_String)
                tb_lotNo.SelectAll()
                tb_lotNo.Focus()
            Else
                tb_Requester.SelectAll()
                tb_Requester.Focus()
            End If
        End If

    End Sub

    Private Sub BTN_Setting_Open_Click(sender As Object, e As EventArgs) Handles BTN_Setting_Open.Click

        If SplitContainer1.Panel1Collapsed = True Then
            SplitContainer1.Panel1Collapsed = False
            BTN_Setting_Open.Text = "프린터 설정 닫기"
        Else
            SplitContainer1.Panel1Collapsed = True
            BTN_Setting_Open.Text = "프린터 설정 열기"
        End If

    End Sub

    Private Sub tb_Requester_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_Requester.KeyDown

        If e.KeyCode = 13 And Not tb_Requester.Text = String.Empty Then
            cb_Process.Focus()
        End If

    End Sub

    Private Sub TB_Port_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Port.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Port", TB_Port.Text)

    End Sub

    Private Sub TB_Width_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_LEFT_Loc.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Width", TB_LEFT_Loc.Text)

    End Sub

    Private Sub TB_Height_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_TOP_Loc.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Height", TB_TOP_Loc.Text)

    End Sub

    Private Sub TB_BaudRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_BaudRate.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "BaudRate", TB_BaudRate.Text)

    End Sub

    Private Sub TB_DataBits_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_DataBits.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "DataBits", TB_DataBits.Text)

    End Sub

    Private Sub TB_Parity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Parity.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Parity", TB_Parity.Text)

    End Sub

    Private Sub TB_StopBits_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_StopBits.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "StopBits", TB_StopBits.Text)

    End Sub

    Private Sub cb_PrinterList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_PrinterList.SelectedIndexChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Printer Name", cb_PrinterList.Text)

    End Sub

    Private Sub cb_Cable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_Cable.SelectedIndexChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "COM/LPT", cb_Cable.Text)

        If cb_Cable.Text = "USB" Then
            PrinterListLoad()
        End If

    End Sub

    Private Sub BTN_SavePrint_Click(sender As Object, e As EventArgs) Handles BTN_SavePrint.Click

        If tb_lotNo.Text = String.Empty Then
            MsgBox("Lot No. 또는 YJ No.가 입력되지 않았습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            tb_lotNo.SelectAll()
            tb_lotNo.Focus()
            Exit Sub
        End If

        If tb_Requester.Text = String.Empty Then
            MsgBox("라벨 발행 요청자가 입력되지 않습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            tb_Requester.SelectAll()
            tb_Requester.Focus()
            Exit Sub
        End If

        If tb_Process_Char.Text = String.Empty Then
            MsgBox("공정이 선택되지 않았습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If cb_Request_Reason.Text = String.Empty Then
            MsgBox("사유가 선택되지 않았습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        Dim printString As String = String.Empty

        If RadioButton1.Checked Then
            printString = RadioButton1.Text
            If tb_Process_No.Text = String.Empty Then
                MsgBox("공정번호가 입력되지 않았습니다.", MsgBoxStyle.Information, form_Msgbox_String)
                Exit Sub
            End If
            If CInt(tb_Process_No.Text) > CInt(tb_total_print_qty.Text) Then
                MsgBox("발행된 공정번호보다 입력한 공정번호가 큽니다.", MsgBoxStyle.Information, form_Msgbox_String)
                Exit Sub
            End If
        ElseIf RadioButton2.Checked Then
            printString = RadioButton2.Text
        End If

        If printString = String.Empty Then
            MsgBox("라벨 발행 방법이 선택되지 않았습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If MsgBox("라벨을 " & printString & "하시겠습니까?",
                      MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Saving...")

        '###### 라벨 발행부분
        Dim swFile As StreamWriter =
            New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))

        Dim lotNo As String = String.Empty
        Dim rcd_information As String = tb_After_RCD.Text & " ( " & tb_After_RCD_Marking.Text & " )"
        Dim pmic_information As String = tb_After_PMIC.Text & " ( " & tb_After_PMIC_Marking.Text & " )"

        'If tb_Before_PMIC.Text = tb_After_PMIC.Text Then
        '    pmic_information = String.Empty 'PMIC 수리하지 않으므로 삭제
        'End If

        'If tb_Before_RCD.Text = tb_After_RCD.Text Then
        '    rcd_information = String.Empty 'RCD 수리하지 않으므로 삭제
        'End If

        If tb_Lot_Option.Text Like "*O" Then '숫자 영 영어 오
            rcd_information = String.Empty
        ElseIf tb_Lot_Option.Text Like "*Q" Then

        Else
            thread_LoadingFormEnd
            Thread.Sleep(100)
            MsgBox("수리 정보가 확실하지 않은 Option 입니다. 확인하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        Dim new_process_no As String = String.Empty

        If RadioButton1.Checked Then
            new_process_no = Format(CInt(tb_Process_No.Text), "00")
        ElseIf RadioButton2.Checked Then
            new_process_no = Format(CInt(tb_total_print_qty.Text) + 1, "00")
        End If

        lotNo = tb_Print_Lot_No.Text & " / " &
            tb_Print_YJ_No.Text & "-" & tb_Process_Char.Text & new_process_no & " / " &
            tb_After_Bucket.Text
        swFile.WriteLine("^XZ~JA^XZ")
        swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
        swFile.WriteLine("^MD15") '진하기
        swFile.WriteLine("^FO000,0000^A0,60,40^FD" & lotNo & "^SF%%%%%%%%%%%%%%%%%%%%%dd,1^FS")
        If tb_Process_Char.Text = "R" Then '폐기자재 회수 라벨이라면...
            If tb_Process_No.Text = "1" Then
                rcd_information = String.Empty
                pmic_information = tb_Before_PMIC.Text & " ( " & tb_Before_PMIC_Marking.Text & " )"
            ElseIf tb_Process_No.Text = "2" Then
                rcd_information = tb_Before_RCD.Text & " ( " & tb_Before_RCD_Marking.Text & " )"
                pmic_information = String.Empty
            End If
            swFile.WriteLine("^FO600,0110^A0,40,40^FD" & nowChipQty & "EA" & "^FS")
        End If
        swFile.WriteLine("^FO000,0060^A0,40,30^FDPMIC : " & pmic_information & "^FS")
        swFile.WriteLine("^FO000,0095^A0,40,30^FDRCD : " & rcd_information & "^FS")
        swFile.WriteLine("^FO000,0135^A0,35,25^FDP/N : " & nowProduct & "^FS")
        swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & Replace(lotNo, " ", "") & "/" & tb_After_PMIC.Text & "/" & tb_After_RCD.Text & "^FS")
        swFile.WriteLine("^PQ1^XZ") 'PQ : 발행수량, "ZX : 종료
        swFile.Close()

        '##### 프린터로 전송하는 부분
        Try
            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", "LPT" & TB_Port.Text)
            ElseIf cb_Cable.Text = "COM" Then
                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While

                sr.Close()
                ComPort.Close()
            ElseIf cb_Cable.Text = "USB" Then
                Dim p As New PrintRaw
                Dim s As New StringBuilder
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    s.AppendLine(str)
                End While

                sr.Close()
                p.Print(s, cb_PrinterList.Text)
            End If
            File.Delete(Application.StartupPath & "\print.txt")
        Catch ex As Exception
            File.Delete(Application.StartupPath & "\print.txt")
            thread_LoadingFormEnd
            Thread.Sleep(100)
            MsgBox("라벨 발행 실패." & vbCrLf & ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try
            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            '라벨 발행 기록
            strSQL += "insert into history_label_print(lot_no, yj_no, process_char, process_no"
            strSQL += ", print_qty, printer, print_date, print_method, print_reason, print_requester) values("
            strSQL += "'" & tb_Print_Lot_No.Text & "'"
            strSQL += ",'" & tb_Print_YJ_No.Text & "'"
            strSQL += ",'" & tb_Process_Char.Text & "'"
            strSQL += ",'" & new_process_no & "'"
            strSQL += ",'1'"
            strSQL += ",'" & loginID & "'"
            strSQL += ",'" & write_date & "'"
            strSQL += ",'" & printString & "'"
            strSQL += ",'" & cb_Request_Reason.Text & "'"
            strSQL += ",'" & tb_Requester.Text & "');"

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            thread_LoadingFormEnd
            Thread.Sleep(100)
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        thread_LoadingFormEnd
        Thread.Sleep(100)
        MsgBox("라벨 발행 완료.", MsgBoxStyle.Information, form_Msgbox_String)

        tb_Print_Lot_No.Text = String.Empty
        tb_Print_YJ_No.Text = String.Empty
        tb_before_Bucket.Text = String.Empty
        tb_After_Bucket.Text = String.Empty
        tb_Before_PMIC.Text = String.Empty
        tb_Before_PMIC_Marking.Text = String.Empty
        tb_After_PMIC.Text = String.Empty
        tb_After_PMIC_Marking.Text = String.Empty
        tb_Before_RCD.Text = String.Empty
        tb_Before_RCD_Marking.Text = String.Empty
        tb_After_RCD.Text = String.Empty
        tb_After_RCD_Marking.Text = String.Empty

        tb_lotNo.Text = String.Empty
        tb_Requester.Text = String.Empty
        tb_Process_No.Text = String.Empty
        cb_Request_Reason.Text = String.Empty
        cb_Process.SelectedIndex = -1
        tb_Process_Char.Text = String.Empty

        history_Load()
        tb_lotNo.SelectAll()
        tb_lotNo.Focus()

    End Sub

    Private Sub PrinterListLoad()

        cb_PrinterList.Items.Clear()

        For i As Integer = 0 To Printing.PrinterSettings.InstalledPrinters.Count - 1
            Dim Printers As String = Printing.PrinterSettings.InstalledPrinters.Item(i)
            cb_PrinterList.Items.Add(Printers)
        Next

        cb_PrinterList.Sorted = True

    End Sub

    Private Sub tb_Process_No_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_Process_No.KeyDown

        Select Case e.KeyCode
            Case 13
            Case 8
            Case 96 To 105, 48 To 57
            Case Else
                e.SuppressKeyPress = True
        End Select

        If Not tb_Process_No.Text = String.Empty Then
            RadioButton1.Checked = True
        End If

    End Sub

    Private Sub tb_lotNo_TextChanged(sender As Object, e As EventArgs) Handles tb_lotNo.TextChanged

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click
        Me.Dispose()
    End Sub

    Private Sub tb_Parts_LotNo_TextChanged(sender As Object, e As EventArgs) Handles tb_Parts_LotNo.TextChanged



    End Sub

    Private Sub tb_Parts_LotNo_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_Parts_LotNo.KeyDown

        If e.KeyCode = 13 And Not tb_Parts_LotNo.Text = String.Empty Then

            tb_Parts_Load_LotNo.Text = String.Empty
            tb_Parts_Load_PartNo.Text = String.Empty
            tb_Parts_Load_TopMarking.Text = String.Empty
            tb_Parts_Load_YJNo.Text = String.Empty
            TextBox1.Text = String.Empty

            DBConnect()

            Try
                Dim strSQL As String = "call sp_label_reprint('" & tb_Parts_LotNo.Text & "', null, 3)"

                Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
                Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

                Do While sqlDR.Read
                    If IsDBNull(sqlDR("pmic_top_marking")) Then
                        TextBox1.Text = "RCD"
                    Else
                        TextBox1.Text = "PMIC"
                    End If
                    tb_Parts_Load_LotNo.Text = sqlDR("lot_no")
                    tb_Parts_Load_PartNo.Text = sqlDR("part_no")
                    If IsDBNull(sqlDR("pmic_top_marking")) Then
                        tb_Parts_Load_TopMarking.Text = sqlDR("rcd_top_marking")
                    Else
                        tb_Parts_Load_TopMarking.Text = sqlDR("pmic_top_marking")
                    End If
                    tb_Parts_Load_YJNo.Text = sqlDR("yj_no")
                Loop
                sqlDR.Close()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            End Try
            DBClose()

            If TextBox1.Text = String.Empty Then
                MsgBox("해당 Lot 정보를 찾을 수 없습니다.", MsgBoxStyle.Information, form_Msgbox_String)
                tb_Parts_LotNo.SelectAll()
                tb_Parts_LotNo.Focus()
            Else
                tb_Parts_Requester.SelectAll()
                tb_Parts_Requester.Focus()
            End If
        End If

    End Sub

    Private Sub btn_Parts_Save_Click(sender As Object, e As EventArgs) Handles btn_Parts_Save.Click

        If tb_Parts_LotNo.Text = String.Empty Then
            MsgBox("Lot No.가 입력되지 않았습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If tb_Parts_Requester.Text = String.Empty Then
            MsgBox("요청가 입력되지 않았습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If cb_Parts_Reason.Text = String.Empty Then
            MsgBox("사유가 선택되지 않았습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If tb_Parts_Load_LotNo.Text = String.Empty Then
            MsgBox("자재 정보가 없습니다..", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        Dim printString As String = "재발행"

        If MsgBox("라벨을 " & printString & "하시겠습니까?",
                      MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart()

        '###### 라벨 발행부분
        Dim swFile As StreamWriter =
            New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))

        swFile.WriteLine("^XZ~JA^XZ")
        swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
        swFile.WriteLine("^MD15") '진하기
        swFile.WriteLine("^FO000,0000^A0,70,55^FDPart No : " & tb_Parts_Load_PartNo.Text & "^FS")
        swFile.WriteLine("^FO000,0065^A0,50,40^FDLot No. : " & tb_Parts_Load_YJNo.Text & " / " & tb_Parts_Load_LotNo.Text & "^FS")
        swFile.WriteLine("^FO000,0115^A0,50,40^FDTop Marking : " & tb_Parts_Load_TopMarking.Text & "^FS")
        swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & tb_Parts_Load_PartNo.Text & "/" & tb_Parts_Load_LotNo.Text & "/" & tb_Parts_Load_YJNo.Text & "^FS")
        swFile.WriteLine("^PQ2^XZ") 'PQ : 발행수량, "ZX : 종료
        swFile.Close()

        '##### 프린터로 전송하는 부분
        Try
            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", "LPT" & TB_Port.Text)
            ElseIf cb_Cable.Text = "COM" Then
                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While

                sr.Close()
                ComPort.Close()
            ElseIf cb_Cable.Text = "USB" Then
                Dim p As New PrintRaw
                Dim s As New StringBuilder
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    s.AppendLine(str)
                End While

                sr.Close()
                p.Print(s, cb_PrinterList.Text)
            End If
            File.Delete(Application.StartupPath & "\print.txt")
        Catch ex As Exception
            File.Delete(Application.StartupPath & "\print.txt")
            thread_LoadingFormEnd
            Thread.Sleep(100)
            MsgBox("라벨 발행 실패." & vbCrLf & ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try
            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            '라벨 발행 기록
            strSQL += "insert into history_label_print(lot_no, yj_no, process_char, process_no"
            strSQL += ", print_qty, printer, print_date, print_method, print_reason, print_requester) values("
            strSQL += "'" & tb_Parts_Load_LotNo.Text & "'"
            strSQL += ",'" & tb_Parts_Load_YJNo.Text & "'"
            strSQL += ",''"
            strSQL += ",'0'"
            strSQL += ",'2'"
            strSQL += ",'" & loginID & "'"
            strSQL += ",'" & write_date & "'"
            strSQL += ",'" & printString & "'"
            strSQL += ",'" & cb_Parts_Reason.Text & "'"
            strSQL += ",'" & tb_Parts_Requester.Text & "');"

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            thread_LoadingFormEnd
            Thread.Sleep(100)
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        thread_LoadingFormEnd
        Thread.Sleep(100)
        MsgBox("라벨 발행 완료.", MsgBoxStyle.Information, form_Msgbox_String)

        tb_Parts_Load_LotNo.Text = String.Empty
        tb_Parts_Load_PartNo.Text = String.Empty
        tb_Parts_Load_TopMarking.Text = String.Empty
        tb_Parts_Load_YJNo.Text = String.Empty
        TextBox1.Text = String.Empty

        tb_Parts_LotNo.Text = String.Empty
        tb_Parts_Requester.Text = String.Empty
        cb_Parts_Reason.Text = String.Empty

        history_Load()

        tb_Parts_LotNo.SelectAll()
        tb_Parts_LotNo.Focus()

    End Sub
End Class
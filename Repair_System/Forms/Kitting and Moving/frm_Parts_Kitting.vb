'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-07-25
'##### 수정일자 : 2023-07-25
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : Baking 완료된 Lot에 필요자재를 매칭시키는 폼

'############################################################################################################
'############################################################################################################

Imports System.IO
Imports System.Text
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_Parts_Kitting

    Dim form_Msgbox_String As String = "Kitting 등록"
    Private WithEvents ComPort As New Ports.SerialPort
    Dim pmic_top_marking As String
    Dim rcd_top_marking As String
    Dim temp_code As String

    Private Sub frm_Parts_Kitting_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        SplitContainer1.Panel1Collapsed = True

        grid_Setting()

        Label13.Visible = False
        Label14.Visible = False

        Setting_LOAD()

    End Sub

    Private Sub grid_Setting()

        With grid_KittingList
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
            grid_KittingList(0, 0) = "No"
            grid_KittingList(0, 1) = "Kitting 일자"
            grid_KittingList(0, 2) = "Lot No."
            grid_KittingList(0, 3) = "YJ No."
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

        With grid_PartsList
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 7
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_PartsList(0, 0) = "No"
            grid_PartsList(0, 1) = "구분"
            grid_PartsList(0, 2) = "Part No."
            grid_PartsList(0, 3) = "Lot No."
            grid_PartsList(0, 4) = "YJ No."
            grid_PartsList(0, 5) = "사용수량"
            grid_PartsList(0, 6) = "확인"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

    End Sub

    Private Sub tb_Barcode_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_Barcode.KeyDown

        If e.KeyCode = 13 And Not tb_Barcode.Text = String.Empty Then

            tb_Barcode.Text = tb_Barcode.Text.ToUpper

            grid_PartsList.Rows.Count = 1
            tb_Product.Text = String.Empty
            tb_Lot_No.Text = String.Empty
            tb_YJ_No.Text = String.Empty
            tb_Lot_Qty.Text = String.Empty
            tb_ORG_Product.Text = String.Empty
            tb_before_CC.Text = String.Empty
            tb_Before_PMIC.Text = String.Empty
            tb_Before_RCD.Text = String.Empty
            tb_After_CC.Text = String.Empty
            tb_After_PMIC.Text = String.Empty
            tb_After_RCD.Text = String.Empty
            tb_Lot_Status.Text = String.Empty

            If UBound(tb_Barcode.Text.Split("/")) < 2 And Not CheckBox1.Checked Then
                MsgBox("Yujin Barcode(Matamatrix)를 스캔하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
                tb_Barcode.SelectAll()
                tb_Barcode.Focus()
                Exit Sub
            Else
                If Not tb_Barcode.Text.Substring(0, 1).Equals("Y") Then
                    MsgBox("Yujin No.를 입력하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
                    tb_Barcode.SelectAll()
                    tb_Barcode.Focus()
                    Exit Sub
                End If
            End If

            Dim use_lot_no As String = tb_Barcode.Text.Split("/")(0)
            If CheckBox1.Checked Then use_lot_no = tb_Barcode.Text
            lot_Information_Load(use_lot_no) 'Lot 정보를 불러온다.

            If tb_Product.Text = String.Empty Then
                MsgBox("Lot 정보를 찾을 수 없습니다.", MsgBoxStyle.Information, form_Msgbox_String)
                tb_Barcode.SelectAll()
                tb_Barcode.Focus()
                Exit Sub
            End If

            If Not tb_Lot_Status.Text = "Baking End" Then
                MsgBox("Baking 완료상태가 아닙니다.", MsgBoxStyle.Information, form_Msgbox_String)
                tb_Barcode.SelectAll()
                tb_Barcode.Focus()
                Exit Sub
            End If

            ''PMIC 정보가 다른지 확인 후 필요자재를 불러온다.
            'If Not tb_Before_PMIC.Text = tb_After_PMIC.Text Then
            '    necessary_Parts(tb_After_PMIC.Text)
            '    Label13.Visible = True
            'Else
            '    Label13.Visible = False
            'End If
            ''RCD 정보가 다른지 확인 후 필요자재를 불러온다.
            'If Not tb_Before_RCD.Text = tb_After_RCD.Text Then
            '    necessary_Parts(tb_After_RCD.Text)
            '    Label14.Visible = True
            'Else
            '    Label14.Visible = False
            'End If

            If tb_Option.Text Like "*O" Then '숫자 영 영어 오
                necessary_Parts(tb_After_PMIC.Text)
                Label13.Visible = True
                Label14.Visible = False
            ElseIf tb_Option.Text Like "*Q" Then
                necessary_Parts(tb_After_PMIC.Text)
                Label13.Visible = True
                necessary_Parts(tb_After_RCD.Text)
                Label13.Visible = True
                Label14.Visible = True
            End If

            tb_Parts_Barcode.SelectAll()
            tb_Parts_Barcode.Focus()
        End If

    End Sub

    Private Sub lot_Information_Load(ByVal lotNo As String)

        DBConnect()

        Try
            Dim strSQL As String = "call sp_parts_kitting(0" &
            ", null" &
            ", null" &
            ", '" & lotNo & "'" &
            ", null" &
            ", null)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                tb_Product.Text = sqlDR("product")
                tb_Lot_No.Text = sqlDR("lot_no")
                tb_YJ_No.Text = sqlDR("yj_no")
                tb_Lot_Qty.Text = sqlDR("chip_qty")
                tb_ORG_Product.Text = sqlDR("org_product")
                tb_before_CC.Text = tb_ORG_Product.Text.Substring(16, 2)
                tb_Before_PMIC.Text = sqlDR("org_pmic_partno")
                tb_Before_RCD.Text = sqlDR("org_rcd_partno")
                tb_After_CC.Text = tb_Product.Text.Substring(16, 2)
                tb_After_PMIC.Text = sqlDR("pmic_partno")
                tb_After_RCD.Text = sqlDR("rcd_partno")
                pmic_top_marking = sqlDR("pmic_top_marking")
                rcd_top_marking = sqlDR("rcd_top_marking")
                tb_Lot_Status.Text = sqlDR("lot_status")
                tb_Option.Text = sqlDR("lot_option")
            Loop
            sqlDR.Close()
        Catch ex As Exception
            Console.WriteLine("Lot 정보 불러오는 중 에러 발생 : " & ex.Message)
        End Try

        DBClose()

    End Sub

    Private Sub necessary_Parts(ByVal part_no As String)

        grid_PartsList.Redraw = False

        DBConnect()

        Dim need_qty As Integer = CInt(tb_Lot_Qty.Text)
        Dim find_qty As Integer = 0

        Dim strSQL As String = "call sp_parts_kitting(1" &
            ", null" &
            ", null" &
            ", '" & part_no & "'" &
            ", null" &
            ", null)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader
        Dim loopStop As Boolean = False

        Do While sqlDR.Read
            Dim available_qty As Integer = sqlDR("available_qty")
            find_qty += available_qty
            If need_qty - find_qty < 1 Then '필요수량이 충족할 경우 탈출
                find_qty = find_qty - available_qty '원래 수량으로 돌리고
                available_qty = need_qty - find_qty '현재 자재lot의 필요수량만!!
                loopStop = True
            End If

            Dim insert_string As String = grid_PartsList.Rows.Count & vbTab &
                                          sqlDR("part_division") & vbTab &
                                          sqlDR("part_no") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          Format(available_qty, "#,##0")
            grid_PartsList.AddItem(insert_string)
            If loopStop = True Then Exit Do
        Loop
        sqlDR.Close()

        DBClose()

        grid_PartsList.AutoSizeCols()
        grid_PartsList.Redraw = True

        If loopStop = False Then
            MsgBox("Part No. : " & part_no & vbCrLf & "재고수량이 부족합니다.", MsgBoxStyle.Information, form_Msgbox_String)
            btn_Save.Enabled = False
            btn_Save2.Enabled = False
        End If

    End Sub

    Private Sub btn_New_Kitting_Click(sender As Object, e As EventArgs) Handles btn_New_Kitting.Click

        temp_code = "PO" & Format(Now, "yyyyMMddHHmmss")

        grid_PartsList.Rows.Count = 1
        tb_Product.Text = String.Empty
        tb_Lot_No.Text = String.Empty
        tb_YJ_No.Text = String.Empty
        tb_Lot_Qty.Text = String.Empty
        tb_ORG_Product.Text = String.Empty
        tb_before_CC.Text = String.Empty
        tb_Before_PMIC.Text = String.Empty
        tb_Before_RCD.Text = String.Empty
        tb_After_CC.Text = String.Empty
        tb_After_PMIC.Text = String.Empty
        tb_After_RCD.Text = String.Empty

        btn_Save.Enabled = True
        btn_Save2.Enabled = True
        tb_Barcode.Enabled = True

        tb_Barcode.SelectAll()
        tb_Barcode.Focus()

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

    Private Sub PrinterListLoad()

        cb_PrinterList.Items.Clear()

        For i As Integer = 0 To Printing.PrinterSettings.InstalledPrinters.Count - 1
            Dim Printers As String = Printing.PrinterSettings.InstalledPrinters.Item(i)
            cb_PrinterList.Items.Add(Printers)
        Next

        cb_PrinterList.Sorted = True

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click

        If SplitContainer1.Panel1Collapsed = True Then
            SplitContainer1.Panel1Collapsed = False
        Else
            SplitContainer1.Panel1Collapsed = True
        End If

    End Sub

    Dim nowPartsQty As Integer

    Private Sub tb_Parts_Barcode_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_Parts_Barcode.KeyDown

        If e.KeyCode = 13 And Not tb_Parts_Barcode.Text = String.Empty Then

            nowPartsQty = 0

            Dim scanPartNo As String = tb_Parts_Barcode.Text.Split("/")(0)
            Dim scanLotNo As String = tb_Parts_Barcode.Text.Split("/")(1)
            Dim scanYJNo As String = tb_Parts_Barcode.Text.Split("/")(2)

            Dim findItem As Boolean = False

            For i = 1 To grid_PartsList.Rows.Count - 1
                If grid_PartsList(i, 2) = scanPartNo And
                        grid_PartsList(i, 3) = scanLotNo And
                        grid_PartsList(i, 4) = scanYJNo Then
                    If grid_PartsList(i, 6) = "OK" Then
                        MsgBox("출고 등록된 자재 입니다.", MsgBoxStyle.Information, form_Msgbox_String)
                        tb_Parts_Barcode.SelectAll()
                        tb_Parts_Barcode.Focus()
                        Exit Sub
                    Else
                        nowPartsQty = CInt(grid_PartsList(i, 5))
                        findItem = True
                    End If
                    Exit For
                End If
            Next

            If findItem = False Then
                MsgBox("List에서 해당 자재를 찾을 수 없습니다." & vbCrLf &
                       "확인하여 주십시오." _
                       , MsgBoxStyle.Information _
                       , form_Msgbox_String)
                tb_Parts_Barcode.SelectAll()
                tb_Parts_Barcode.Focus()
                Exit Sub
            Else
                tb_Parts_PartNo.Text = scanPartNo
                tb_Parts_Lot_No.Text = scanLotNo
                tb_Parts_YJ_No.Text = scanYJNo
                tb_Parts_Qty.SelectAll()
                tb_Parts_Qty.Focus()
            End If
        End If

    End Sub

    Private Sub tb_Parts_Qty_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_Parts_Qty.KeyDown

        Select Case e.KeyCode
            Case 13
            Case 8
            Case 96 To 105, 48 To 57
            Case Else
                e.SuppressKeyPress = True
        End Select

        If e.KeyCode = 13 And Not tb_Parts_Qty.Text = String.Empty Then
            If nowPartsQty = CInt(tb_Parts_Qty.Text) Then
                If MsgBox("확인 되었습니다." & vbCrLf &
                          "바코드 라벨을 출력 하시겠습니까" _
                          , MsgBoxStyle.Question + MsgBoxStyle.YesNo _
                          , form_Msgbox_String) = MsgBoxResult.No Then Exit Sub
                btn_Check_Print_Click(Nothing, Nothing)
            Else
                MsgBox("출고 수량을 확인하여 주십시오." _
                       , MsgBoxStyle.Information _
                       , form_Msgbox_String)
                tb_Parts_Qty.SelectAll()
                tb_Parts_Qty.Focus()
            End If
        End If

    End Sub

    Private Sub btn_Check_Print_Click(sender As Object, e As EventArgs) Handles btn_Check_Print.Click

        If tb_Parts_PartNo.Text = String.Empty Then
            MsgBox("확인 자재가 선택되지 않았습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            tb_Parts_Barcode.SelectAll()
            tb_Parts_Barcode.Focus()
            Exit Sub
        End If

        Dim now_top_marking As String = String.Empty
        If tb_After_PMIC.Text = tb_Parts_PartNo.Text Then
            now_top_marking = pmic_top_marking
        ElseIf tb_After_RCD.Text = tb_Parts_PartNo.Text Then
            now_top_marking = rcd_top_marking
        End If

        '###### 라벨 발행부분
        Dim swFile As StreamWriter =
            New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))

        Dim datamatrix_Data As String = tb_Parts_PartNo.Text & "/" &
            tb_Parts_Lot_No.Text & "/" &
            tb_Parts_YJ_No.Text & "/" &
            tb_YJ_No.Text

        swFile.WriteLine("^XZ~JA^XZ")
        swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
        swFile.WriteLine("^MD15") '진하기
        swFile.WriteLine("^FO000,0000^A0,70,55^FDPart No : " & tb_Parts_PartNo.Text & "^FS")
        swFile.WriteLine("^FO000,0065^A0,50,40^FDTop Marking : " & now_top_marking & "^FS")
        swFile.WriteLine("^FO000,0115^A0,50,40^FDLot No. : " & tb_Parts_Lot_No.Text & "/" & tb_Parts_YJ_No.Text & "^FS")
        swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & datamatrix_Data & ")^FS")
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
            MsgBox("라벨 발행 실패." & vbCrLf & ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try


        Dim scanPartNo As String = tb_Parts_Barcode.Text.Split("/")(0)
        Dim scanLotNo As String = tb_Parts_Barcode.Text.Split("/")(1)
        Dim scanYJNo As String = tb_Parts_Barcode.Text.Split("/")(2)

        For i = 1 To grid_PartsList.Rows.Count - 1
            If grid_PartsList(i, 2) = scanPartNo And
                    grid_PartsList(i, 3) = scanLotNo And
                    grid_PartsList(i, 4) = scanYJNo Then
                grid_PartsList(i, 6) = "OK"
                Exit For
            End If
        Next

        tb_Parts_Barcode.Text = String.Empty
        tb_Parts_PartNo.Text = String.Empty
        tb_Parts_Lot_No.Text = String.Empty
        tb_Parts_YJ_No.Text = String.Empty
        tb_Parts_Qty.Text = String.Empty
        tb_Parts_Barcode.SelectAll()
        tb_Parts_Barcode.Focus()

        If allOK_Check() = True Then
            MsgBox("모든 출고 리스트의 자재를 확인 하였습니다." & vbCrLf &
                   "Kitting 정보를 저장하여 주십시오." _
                   , MsgBoxStyle.Information _
                   , form_Msgbox_String)
        End If
    End Sub

    Private Function allOK_Check() As Boolean

        allOK_Check = False

        Dim okCount As Integer = 0

        For i = 1 To grid_PartsList.Rows.Count - 1
            If grid_PartsList(i, 6) = "OK" Then
                okCount += 1
            End If
        Next

        If okCount = grid_PartsList.Rows.Count - 1 Then
            allOK_Check = True
        End If

        Return allOK_Check

    End Function

    Private Sub btn_Save2_Click(sender As Object, e As EventArgs) Handles btn_Save2.Click, btn_Save.Click

        If allOK_Check() = False Then
            MsgBox("모든 출고 리스트가 확인되지 않았습니다.." _
                   , MsgBoxStyle.Information _
                   , form_Msgbox_String)
            Exit Sub
        End If

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try
            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            '라벨 발행 기록
            For i = 1 To grid_PartsList.Rows.Count - 1
                strSQL += "insert into kitting_material_information(temp_code, part_division, part_no, lot_no, yj_no"
                strSQL += ", used_lot_no, used_yj_no, kitting_qty, worker, write_date, first_write_date) values("
                strSQL += "'" & temp_code & i & "'"
                strSQL += ",'" & grid_PartsList(i, 1) & "'"
                strSQL += ",'" & grid_PartsList(i, 2) & "'"
                strSQL += ",'" & grid_PartsList(i, 3) & "'"
                strSQL += ",'" & grid_PartsList(i, 4) & "'"
                strSQL += ",'" & tb_Lot_No.Text & "'"
                strSQL += ",'" & tb_YJ_No.Text & "'"
                strSQL += ",'" & grid_PartsList(i, 5) & "'"
                strSQL += ",'" & loginID & "'"
                strSQL += ",'" & write_date & "'"
                strSQL += ",'" & write_date & "');"
            Next

            strSQL += "update basic_lot_information set lot_status = 'Kitting Completed'"
            strSQL += " where lot_no = '" & tb_Lot_No.Text & "';"

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        MsgBox("Kitting 정보 등록을 완료 하였습니다.", MsgBoxStyle.Information, form_Msgbox_String)

        grid_PartsList.Rows.Count = 1
        tb_Product.Text = String.Empty
        tb_Lot_No.Text = String.Empty
        tb_YJ_No.Text = String.Empty
        tb_Lot_Qty.Text = String.Empty
        tb_ORG_Product.Text = String.Empty
        tb_before_CC.Text = String.Empty
        tb_Before_PMIC.Text = String.Empty
        tb_Before_RCD.Text = String.Empty
        tb_After_CC.Text = String.Empty
        tb_After_PMIC.Text = String.Empty
        tb_After_RCD.Text = String.Empty
        tb_Lot_Status.Text = String.Empty

        tb_Parts_Barcode.Text = String.Empty
        tb_Parts_PartNo.Text = String.Empty
        tb_Parts_Lot_No.Text = String.Empty
        tb_Parts_YJ_No.Text = String.Empty
        tb_Parts_Qty.Text = String.Empty

        tb_Barcode.SelectAll()
        tb_Barcode.Focus()

    End Sub

    Private Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        btn_Save.Enabled = False
        btn_Save2.Enabled = False
        tb_Barcode.Enabled = False

        grid_KittingList.Rows.Count = 1
        grid_KittingList.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_parts_kitting(2" &
            ", '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "'" &
            ", '" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker2.Value), "yyyy-MM-dd") & "'" &
            ", '" & tb_S_YJ_No.Text.ToUpper & "'" &
            ", '" & tb_S_Lot_No.Text.ToUpper & "'" &
            ", null)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_string As String = grid_KittingList.Rows.Count & vbTab &
                                          Format(sqlDR("first_write_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          sqlDR("used_lot_no") & vbTab &
                                          sqlDR("used_yj_no")
            grid_KittingList.AddItem(insert_string)
        Loop
        sqlDR.Close()

        DBClose()

        grid_KittingList.AutoSizeCols()
        grid_KittingList.Redraw = True

    End Sub

    Private Sub grid_KittingList_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles grid_KittingList.MouseDoubleClick

        If e.Button = MouseButtons.Left And grid_KittingList.MouseRow > 0 Then
            Dim selRow As Integer = grid_KittingList.MouseRow
            Dim sel_LotNo As String = grid_KittingList(selRow, 2)
            lot_Information_Load(sel_LotNo)

            grid_PartsList.Rows.Count = 1
            grid_PartsList.Redraw = False

            DBConnect()

            Dim strSQL As String = "call sp_parts_kitting(3" &
                ", null" &
                ", null" &
                ", '" & sel_LotNo & "'" &
                ", null" &
                ", null)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                Dim insert_string As String = grid_PartsList.Rows.Count & vbTab &
                                              sqlDR("part_division") & vbTab &
                                              sqlDR("part_no") & vbTab &
                                              sqlDR("lot_no") & vbTab &
                                              sqlDR("yj_no") & vbTab &
                                              Format(sqlDR("kitting_qty"), "#,##0")
                grid_PartsList.AddItem(insert_string)
            Loop
            sqlDR.Close()

            DBClose()

            grid_PartsList.AutoSizeCols()
            grid_PartsList.Redraw = True
        End If

    End Sub

    Private Sub tb_Barcode_TextChanged(sender As Object, e As EventArgs) Handles tb_Barcode.TextChanged

    End Sub

    Private Sub tb_Barcode_Layout(sender As Object, e As LayoutEventArgs) Handles tb_Barcode.Layout

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click
        Me.Dispose()
    End Sub

    Private Sub tb_Parts_Qty_TextChanged(sender As Object, e As EventArgs) Handles tb_Parts_Qty.TextChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        tb_Barcode.SelectAll()
        tb_Barcode.Focus()

    End Sub
End Class
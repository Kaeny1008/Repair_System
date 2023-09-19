'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-07-12
'##### 수정일자 : 2023-07-12
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 신규 입고되는 자재를 등록한다.
'                 (OMS 데이터 2종과 우익에서 보내주는 자료를 토대로 입고를 잡는방법)

'############################################################################################################
'############################################################################################################

Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient
Imports System.Threading

Public Class frm_Module_In_Add

    Dim form_Msgbox_String As String = "자재 입고등록"
    Dim runProcess As Thread

    Private Sub MODEL_MNG_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        dtp_Start.Value = Format(Now, "yyyy-MM-01 HH:mm:ss")
        dtp_End.Value = Now

        GRID_SETTING()

    End Sub

    Private Sub GRID_SETTING()

        With grid_SlipNo
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
            grid_SlipNo(0, 0) = "No"
            grid_SlipNo(0, 1) = "Slip No."
            grid_SlipNo(0, 2) = "저장일시"
            grid_SlipNo(0, 3) = "확정시간"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

        With grid_SlipNoDetail
            .AllowEditing = True
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 20
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_SlipNoDetail(0, 0) = "No"
            grid_SlipNoDetail(0, 1) = "구분"
            grid_SlipNoDetail(0, 2) = "Slip No."
            grid_SlipNoDetail(0, 3) = "Product Code"
            grid_SlipNoDetail(0, 4) = "Lot No."
            grid_SlipNoDetail(0, 5) = "Lot Type"
            grid_SlipNoDetail(0, 6) = "Wafer Qty"
            grid_SlipNoDetail(0, 7) = "Chip Qty"
            grid_SlipNoDetail(0, 8) = "Confirm"
            grid_SlipNoDetail(0, 9) = "주차"
            grid_SlipNoDetail(0, 10) = "Step"
            grid_SlipNoDetail(0, 11) = "Issue Date"
            grid_SlipNoDetail(0, 12) = "등록날짜"
            grid_SlipNoDetail(0, 13) = "Confirm Date"
            grid_SlipNoDetail(0, 14) = "Fab Line"
            grid_SlipNoDetail(0, 15) = "Return Type"
            grid_SlipNoDetail(0, 16) = "Option"
            grid_SlipNoDetail(0, 17) = "PFQ, DOE 물량" & vbCrLf & "Or Dummy"
            grid_SlipNoDetail(0, 18) = "비고"
            grid_SlipNoDetail(0, 19) = "Temp Code"
            .Cols(17).DataType = GetType(Boolean)
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 2).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .ExtendLastCol = True
            .Cols(.Cols.Count - 1).Visible = False
            .AutoSizeCols()
            .ShowCursor = True
        End With

    End Sub

    Private Sub BTN_File_Load1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_File_Load1.Click

        grid_SlipNoDetail.Rows.Count = 1

        If Not frm_M_First_Popup.Visible Then frm_M_First_Popup.Show()
        frm_M_First_Popup.Focus()

    End Sub

    Dim open_Cancel As Boolean
    Public open_File_Path As String
    Public slip_no, product_code, lot_no, lot_type, wafer_qty, chip_qty, confirm, ww, issue_date, reg_date, confirm_date, start_row, work_step As Integer
    Public return_type, fab_line, lot_option As Integer

    Private Delegate Sub ButtonEnabledDelegate(ByVal buttonName As Button, ByVal status As Boolean)
    Private ButtonEnabled = New ButtonEnabledDelegate(AddressOf ButtonStatus)

    Private Sub ButtonStatus(ByVal ButtonName As Button, ByVal Status As Boolean)

        ButtonName.Enabled = Status

    End Sub

    Private Delegate Sub SetPartsListDelegate(ByVal addText As String)
    Private SetPartsList As SetPartsListDelegate = New SetPartsListDelegate(AddressOf GridROW)

    Private Sub GridROW(ByVal AddText As String)

        grid_SlipNoDetail.AddItem(AddText)
        grid_SlipNoDetail.AutoSizeCols()
        grid_SlipNoDetail.AutoSizeRows()
        grid_SlipNoDetail.Rows(0).Height = 30

    End Sub

    Private Delegate Sub sumTotalQtyDelegate()
    Private SetsumTotalQty As sumTotalQtyDelegate = New sumTotalQtyDelegate(AddressOf sumTotalQty)

    Private Sub sumTotalQty()

        Dim waferQty As Integer = 0
        Dim chipQty As Integer = 0

        For i = 1 To grid_SlipNoDetail.Rows.Count - 1
            waferQty += CInt(grid_SlipNoDetail(i, 6))
            chipQty += CInt(grid_SlipNoDetail(i, 7))
        Next

        'grid_SlipNoDetail.AddItem()

    End Sub

    Private Delegate Sub SetPGStatusDelegate(ByVal statusText As String)
    Private SetPGStatus As SetPGStatusDelegate = New SetPGStatusDelegate(AddressOf PG_Status)

    Private Sub PG_Status(ByVal statusText As String)

        frm_Main.lb_Status.Text = statusText

    End Sub

    Private Delegate Sub MsgBoxDelegate(ByVal ShowText As String)
    Private MsgboxRun = New MsgBoxDelegate(AddressOf MsgShow)

    Private Sub MsgShow(ByVal ShowText As String)

        MsgBox(ShowText, MsgBoxStyle.Information, form_Msgbox_String)

    End Sub

    Private Delegate Sub GridRedrawDelegate(ByVal status As Boolean)
    Private SetGridRedraw = New GridRedrawDelegate(AddressOf GridRedraw)

    Private Sub GridRedraw(ByVal status As Boolean)

        grid_SlipNoDetail.Redraw = status

    End Sub

    Private Delegate Sub RTUpdateDelegate(ByVal up_lot_no As String, ByVal up_return_type As String, ByVal up_confirm_date As String)
    Private SetRTUpdate = New RTUpdateDelegate(AddressOf RTUpdate)

    Private Sub RTUpdate(ByVal up_lot_no As String, ByVal up_return_type As String, ByVal up_confirm_date As String)

        Dim find_row As Integer = grid_SlipNoDetail.FindRow(up_lot_no, 1, 4, True, True, True)

        If find_row < 1 Then Exit Sub

        If Not up_confirm_date = String.Empty Then
            If grid_SlipNoDetail(find_row, 8) = "NO" Then
                grid_SlipNoDetail(find_row, 8) = "OK"
                If Not grid_SlipNoDetail(find_row, 0) = "N" Then
                    grid_SlipNoDetail(find_row, 0) = "M"
                    grid_SlipNoDetail.Rows(find_row).StyleNew.ForeColor = Color.Red
                End If
            End If

            If Trim(grid_SlipNoDetail(find_row, 13)) = String.Empty Then
                grid_SlipNoDetail(find_row, 13) = up_confirm_date
                If Not grid_SlipNoDetail(find_row, 0) = "N" Then
                    grid_SlipNoDetail(find_row, 0) = "M"
                    grid_SlipNoDetail.Rows(find_row).StyleNew.ForeColor = Color.Red
                End If
            End If
        End If

        '기존 Return Type과 신규 Return Type이 같지 않고 신규 Return Type이 비어 있지 않다면
        If Not grid_SlipNoDetail(find_row, 15) = up_return_type Then
            grid_SlipNoDetail(find_row, 15) = up_return_type
            If Not grid_SlipNoDetail(find_row, 0) = "N" Then
                grid_SlipNoDetail(find_row, 0) = "M"
                grid_SlipNoDetail.Rows(find_row).StyleNew.ForeColor = Color.Red
            End If
        End If

        grid_SlipNoDetail.AutoSizeCols()

    End Sub

    Private Delegate Sub FabUpdateDelegate(ByVal up_lot_no As String, ByVal up_fab_line As String, ByVal up_lot_option As String)
    Private SetFabUpdate = New FabUpdateDelegate(AddressOf FabUpdate)

    Private Sub FabUpdate(ByVal up_lot_no As String, ByVal up_fab_line As String, ByVal up_lot_option As String)

        Dim find_row As Integer = grid_SlipNoDetail.FindRow(up_lot_no, 1, 4, True, True, True)

        If find_row < 1 Then Exit Sub

        If grid_SlipNoDetail(find_row, 14) = String.Empty Then
            grid_SlipNoDetail(find_row, 14) = up_fab_line
            If Not grid_SlipNoDetail(find_row, 0) = "N" Then
                grid_SlipNoDetail(find_row, 0) = "M"
                grid_SlipNoDetail.Rows(find_row).StyleNew.ForeColor = Color.Red
            End If
        End If

        If grid_SlipNoDetail(find_row, 16) = String.Empty Then
            grid_SlipNoDetail(find_row, 16) = up_lot_option
            If Not grid_SlipNoDetail(find_row, 0) = "N" Then
                grid_SlipNoDetail(find_row, 0) = "M"
                grid_SlipNoDetail.Rows(find_row).StyleNew.ForeColor = Color.Red
            End If
        End If

        grid_SlipNoDetail.AutoSizeCols()

    End Sub

    Private Delegate Sub NewLineDelegate()
    Private SetNewLine = New NewLineDelegate(AddressOf GRID_NewLine)

    Private Sub GRID_NewLine()

        For i = 1 To grid_SlipNoDetail.Rows.Count - 1
            If grid_SlipNoDetail(i, 0).ToString = "N" Then
                grid_SlipNoDetail.Rows(i).StyleNew.ForeColor = Color.Blue
            End If
        Next

    End Sub

    Public Sub File_Open_Start1()

        runProcess = New Thread(AddressOf File_Open1)
        runProcess.IsBackground = True
        runProcess.SetApartmentState(ApartmentState.STA) 'OpenFileDialog를 사용하기위해선 STA로 해야되던데...
        runProcess.Start()

    End Sub

    Private Sub File_Open1()

        Invoke(SetGridRedraw, False)
        open_Cancel = False

        Invoke(ButtonEnabled, BTN_File_Load1, False)
        Invoke(ButtonEnabled, BTN_File_Load2, False)
        Invoke(ButtonEnabled, BTN_File_Load3, False)

        Dim excel_App As Object

        excel_App = CreateObject("Excel.Application")
        excel_App.WorkBooks.Open(open_File_Path)

        Invoke(SetPGStatus, "Excel 파일을 열었습니다.") '현재 진행상태를 표시

        excel_App.Visible = False

        Dim sheet_no As Integer = 1 'Open 하려는 시트 순번을 입력, 시트이름을 입력해도 된다.(우익 자료같은경우)
        Dim write_code As Double = Format(Now, "yyMMddHHmmss") * 100

        With excel_App.ActiveWorkbook.Sheets(sheet_no)
            For i = start_row To .UsedRange.Rows.Count - 1
                If open_Cancel = True Then '불러오기 중단을 요청한 경우
                    Invoke(MsgboxRun, "데이터 불러오기를 중단 하였습니다.")
                    Invoke(SetPGStatus, "데이터 불러오기를 중단 하였습니다.     " & i - start_row + 1 & " / " & .UsedRange.Rows.Count - start_row & " 행") '현재 진행상태를 표시
                    Invoke(ButtonEnabled, BTN_File_Load1, True)
                    Invoke(ButtonEnabled, BTN_File_Load2, True)
                    Invoke(ButtonEnabled, BTN_File_Load3, True)
                    excel_App.WorkBooks(1).Close()
                    excel_App.Quit()
                    excel_App = Nothing
                    runProcess.Abort()
                End If

                Invoke(SetPGStatus, "데이터를 불러오고 있습니다.     " & i - start_row + 1 & " / " & .UsedRange.Rows.Count - start_row & " 행") '현재 진행상태를 표시

                Dim r_slip_no As String = .Cells(i, slip_no).Value
                Dim r_product_code As String = .Cells(i, product_code).Value
                Dim r_lot_no As String = .Cells(i, lot_no).Value
                Dim r_lot_type As String = .Cells(i, lot_type).Value
                Dim r_wafer_qty As String = .Cells(i, wafer_qty).Value
                Dim r_chip_qty As String = .Cells(i, chip_qty).Value
                Dim r_confirm As String = .Cells(i, confirm).Value
                Dim r_ww As String = .Cells(i, ww).Value
                Dim r_step As String = .Cells(i, work_step).value
                Dim r_issue_date As String = .Cells(i, issue_date).Value
                Dim r_reg_date As String = .Cells(i, reg_date).Value
                Dim r_confirm_date As String = .Cells(i, confirm_date).Value
                Dim item_division As String = String.Empty

                If r_product_code Like "M*" Then
                    item_division = "Module"
                ElseIf r_product_code Like "K*" Then
                    item_division = "Component"
                End If

                write_code += 1

                Try
                    Invoke(SetPartsList, "N" & vbTab &
                                         item_division & vbTab &
                                         r_slip_no & vbTab &
                                         r_product_code & vbTab &
                                         r_lot_no & vbTab &
                                         r_lot_type & vbTab &
                                         Format(CInt(r_wafer_qty), "#,##0") & vbTab &
                                         Format(CInt(r_chip_qty), "#,##0") & vbTab &
                                         r_confirm & vbTab &
                                         r_ww & vbTab &
                                         r_step & vbTab &
                                         r_issue_date & vbTab &
                                         r_reg_date & vbTab &
                                         r_confirm_date & vbTab &
                                         String.Empty & vbTab &
                                         String.Empty & vbTab &
                                         String.Empty & vbTab &
                                         "False" & vbTab &
                                         String.Empty & vbTab &
                                         "MCIN" & write_code)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
                    excel_App.WorkBooks(1).Close()
                    excel_App.Quit()
                    excel_App = Nothing

                    Invoke(ButtonEnabled, BTN_File_Load1, True)
                    Invoke(ButtonEnabled, BTN_File_Load2, True)
                    Invoke(ButtonEnabled, BTN_File_Load3, True)
                    runProcess.Abort()
                End Try

            Next
        End With

        excel_App.WorkBooks(1).Close()
        excel_App.Quit()
        excel_App = Nothing

        Invoke(SetPGStatus, "성공적으로 데이터를 불러왔습니다.") '현재 진행상태를 표시

        Invoke(ButtonEnabled, BTN_File_Load1, True)
        Invoke(ButtonEnabled, BTN_File_Load2, True)
        Invoke(ButtonEnabled, BTN_File_Load3, True)
        Invoke(SetGridRedraw, True)

        Invoke(SetNewLine)

        Invoke(MsgboxRun, "완료 하였습니다." & vbCrLf & "데이터를 확인하여 주십시오.")

        runProcess.Abort()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        open_Cancel = True

    End Sub

    Private Sub BTN_File_Load2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_File_Load2.Click

        If Not frm_M_Second_Popup.Visible Then frm_M_Second_Popup.Show()
        frm_M_Second_Popup.Focus()

    End Sub

    Public Sub File_Open_Start2()

        runProcess = New Thread(AddressOf File_Open2)
        runProcess.IsBackground = True
        runProcess.SetApartmentState(ApartmentState.STA) 'OpenFileDialog를 사용하기위해선 STA로 해야되던데...
        runProcess.Start()

    End Sub

    Private Sub File_Open2()

        Invoke(SetGridRedraw, False)
        open_Cancel = False

        Invoke(ButtonEnabled, BTN_File_Load1, False)
        Invoke(ButtonEnabled, BTN_File_Load2, False)
        Invoke(ButtonEnabled, BTN_File_Load3, False)

        Dim excel_App As Object

        excel_App = CreateObject("Excel.Application")
        excel_App.WorkBooks.Open(open_File_Path)

        Invoke(SetPGStatus, "Excel 파일을 열었습니다.") '현재 진행상태를 표시

        excel_App.Visible = False

        Dim sheet_no As Integer = 1 'Open 하려는 시트 순번을 입력, 시트이름을 입력해도 된다.(우익 자료같은경우)

        With excel_App.ActiveWorkbook.Sheets(sheet_no)
            For i = start_row To .UsedRange.Rows.Count - 2 '현재 시트에서 마지막행은 빈데이터라 -2를 한다.
                If open_Cancel = True Then '불러오기 중단을 요청한 경우
                    Invoke(MsgboxRun, "데이터 불러오기를 중단 하였습니다.")
                    Invoke(SetPGStatus, "데이터 불러오기를 중단 하였습니다.     " & i - start_row + 1 & " / " & .UsedRange.Rows.Count - 1 - start_row & " 행") '현재 진행상태를 표시
                    Invoke(ButtonEnabled, BTN_File_Load1, True)
                    Invoke(ButtonEnabled, BTN_File_Load2, True)
                    Invoke(ButtonEnabled, BTN_File_Load3, True)
                    excel_App.WorkBooks(1).Close()
                    excel_App.Quit()
                    excel_App = Nothing
                    runProcess.Abort()
                End If

                Invoke(SetPGStatus, "데이터를 불러오고 있습니다.     " & i - start_row + 1 & " / " & .UsedRange.Rows.Count - 1 - start_row & " 행") '현재 진행상태를 표시

                Dim r_lot_no As String = .Cells(i, lot_no).Value
                Dim r_return_type As String = .Cells(i, return_type).Value
                Dim r_confirm_date As String = .Cells(i, confirm_date).Value

                Invoke(SetRTUpdate, r_lot_no, r_return_type, r_confirm_date)

            Next
        End With

        excel_App.WorkBooks(1).Close()
        excel_App.Quit()
        excel_App = Nothing

        Invoke(SetPGStatus, "성공적으로 데이터를 불러왔습니다.") '현재 진행상태를 표시

        Invoke(ButtonEnabled, BTN_File_Load1, True)
        Invoke(ButtonEnabled, BTN_File_Load2, True)
        Invoke(ButtonEnabled, BTN_File_Load3, True)
        Invoke(SetGridRedraw, True)

        Invoke(MsgboxRun, "완료 하였습니다." & vbCrLf & "데이터를 확인하여 주십시오.")

        runProcess.Abort()

    End Sub

    Private Sub BTN_File_Load3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_File_Load3.Click

        If Not frm_M_Third_Popup.Visible Then frm_M_Third_Popup.Show()
        frm_M_Third_Popup.Focus()

    End Sub

    Public Sub File_Open_Start3()

        runProcess = New Thread(AddressOf File_Open3)
        runProcess.IsBackground = True
        runProcess.SetApartmentState(ApartmentState.STA) 'OpenFileDialog를 사용하기위해선 STA로 해야되던데...
        runProcess.Start()

    End Sub

    Private Sub File_Open3()

        Invoke(SetGridRedraw, False)
        open_Cancel = False

        Invoke(ButtonEnabled, BTN_File_Load1, False)
        Invoke(ButtonEnabled, BTN_File_Load2, False)
        Invoke(ButtonEnabled, BTN_File_Load3, False)

        Dim excel_App As Object

        excel_App = CreateObject("Excel.Application")
        excel_App.WorkBooks.Open(open_File_Path)

        Invoke(SetPGStatus, "Excel 파일을 열었습니다.") '현재 진행상태를 표시

        excel_App.Visible = False

        Dim sheet_no As Integer = 1 'Open 하려는 시트 순번을 입력, 시트이름을 입력해도 된다.(우익 자료같은경우)

        With excel_App.ActiveWorkbook.Sheets(sheet_no)

            'Console.WriteLine(start_row)
            Console.WriteLine(.UsedRange.Rows.Count)

            For i = start_row To .UsedRange.Rows.Count + 1
                If open_Cancel = True Then '불러오기 중단을 요청한 경우
                    Invoke(MsgboxRun, "데이터 불러오기를 중단 하였습니다.")
                    Invoke(SetPGStatus, "데이터 불러오기를 중단 하였습니다.     " & i - start_row + 1 & " / " & .UsedRange.Rows.Count + 2 - start_row & " 행") '현재 진행상태를 표시
                    Invoke(ButtonEnabled, BTN_File_Load1, True)
                    Invoke(ButtonEnabled, BTN_File_Load2, True)
                    Invoke(ButtonEnabled, BTN_File_Load3, True)
                    excel_App.WorkBooks(1).Close()
                    excel_App.Quit()
                    excel_App = Nothing
                    runProcess.Abort()
                End If

                Invoke(SetPGStatus, "데이터를 불러오고 있습니다.     " & i - start_row + 1 & " / " & .UsedRange.Rows.Count + 2 - start_row & " 행") '현재 진행상태를 표시

                Dim r_lot_no As String = .Cells(i, lot_no).Value
                Dim r_fab_line As String = .Cells(i, fab_line).Value
                Dim r_lot_option As String = .Cells(i, lot_option).Value

                Invoke(SetFabUpdate, r_lot_no, r_fab_line, r_lot_option)

            Next
        End With

        excel_App.WorkBooks(1).Close()
        excel_App.Quit()
        excel_App = Nothing

        Invoke(SetPGStatus, "성공적으로 데이터를 불러왔습니다.") '현재 진행상태를 표시

        Invoke(ButtonEnabled, BTN_File_Load1, True)
        Invoke(ButtonEnabled, BTN_File_Load2, True)
        Invoke(ButtonEnabled, BTN_File_Load3, True)
        Invoke(SetGridRedraw, True)

        Invoke(MsgboxRun, "완료 하였습니다." & vbCrLf & "데이터를 확인하여 주십시오.")

        runProcess.Abort()

    End Sub

    Private Sub BTN_New_Model_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_New_SlipNo.Click

        grid_SlipNoDetail.Rows.Count = 1
        BTN_File_Load1.Enabled = True
        BTN_File_Load2.Enabled = True
        BTN_File_Load3.Enabled = True
        btn_Save.Enabled = True
        btn_Save2.Enabled = True

        frm_Main.lb_Status.Text = "입고 확정전 데이터를 입력시 차후 확정 후 '2. Return Type 불러오기 (OMS)'를 재 진행하여 주십시오."

        MsgBox("입고 확정전 데이터를 입력시 차후 확정 후" & vbCrLf &
               "'2. Return Type 불러오기 (OMS)'를 재 진행하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)

    End Sub

    Private Sub BTN_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save.Click, btn_Save2.Click

        If MsgBox("저장 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Saving...")

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try

            '##### 입고 리스트를 저장한다. 시작 #####

            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            For i = 1 To grid_SlipNoDetail.Rows.Count - 1
                If grid_SlipNoDetail(i, 0).ToString = "N" Then
                    '현재행이 신규 입력행일 경우
                    strSQL += "insert into basic_lot_information(part_division, slip_no, product, lot_no, lot_type, wafer_qty, chip_qty"
                    strSQL += ", confirm, lot_ww, step, issue_date, reg_date, confirm_date, fab_line, return_type, lot_option"
                    strSQL += ", yj_no, lot_status, writer, write_date, note, first_write_date, pfq_doe, temp_code) values"
                    strSQL += "('" & grid_SlipNoDetail(i, 1) & "'" 'Slip No
                    strSQL += ",'" & grid_SlipNoDetail(i, 2) & "'" 'Slip No
                    strSQL += ",'" & grid_SlipNoDetail(i, 3) & "'" 'Product No.
                    strSQL += ",'" & grid_SlipNoDetail(i, 4) & "'" 'Lot No.
                    strSQL += ",'" & grid_SlipNoDetail(i, 5) & "'" ' Lot Type
                    strSQL += ",'" & CInt(grid_SlipNoDetail(i, 6)) & "'" 'Wafer Qty
                    strSQL += ",'" & CInt(grid_SlipNoDetail(i, 7)) & "'" 'Chip Qty
                    strSQL += ",'" & grid_SlipNoDetail(i, 8) & "'" 'Confirm
                    strSQL += ",'" & grid_SlipNoDetail(i, 9) & "'" 'WW 
                    strSQL += ",'" & grid_SlipNoDetail(i, 10) & "'" 'Step
                    strSQL += ",'" & grid_SlipNoDetail(i, 11) & "'" 'Issue Date
                    strSQL += ",'" & grid_SlipNoDetail(i, 12) & "'" '등록날짜
                    If trim(grid_SlipNoDetail(i, 13)) = String.Empty Then 'Confirm Date
                        strSQL += ", null"
                    Else
                        strSQL += ",'" & grid_SlipNoDetail(i, 13) & "'"
                    End If
                    strSQL += ",'" & grid_SlipNoDetail(i, 14) & "'" 'Fab Line
                    strSQL += ",'" & grid_SlipNoDetail(i, 15) & "'" 'Return Type
                    strSQL += ",'" & grid_SlipNoDetail(i, 16) & "'" 'Option
                    strSQL += ", null"
                    If grid_SlipNoDetail(i, 8) = "NO" Then
                        strSQL += ", 'Confirm Ready'"
                    Else
                        strSQL += ", 'Ready'"
                    End If
                    strSQL += ",'" & loginID & "'"
                    strSQL += ",'" & write_date & "'"
                    strSQL += ",'" & grid_SlipNoDetail(i, 18) & "'" '비고
                    strSQL += ",'" & write_date & "'"
                    strSQL += ",'" & grid_SlipNoDetail(i, 17) & "'" '비고
                    strSQL += ",'" & grid_SlipNoDetail(i, 19) & "');"
                ElseIf grid_SlipNoDetail(i, 0).ToString = "M" Then
                    '현재행이 수정된 경우
                    strSQL += "update basic_lot_information set slip_no = '" & grid_SlipNoDetail(i, 2) & "'"
                    strSQL += ", product = '" & grid_SlipNoDetail(i, 3) & "'"
                    strSQL += ", lot_no = '" & grid_SlipNoDetail(i, 4) & "'"
                    strSQL += ", lot_type = '" & grid_SlipNoDetail(i, 5) & "'"
                    strSQL += ", wafer_qty = '" & CInt(grid_SlipNoDetail(i, 6)) & "'"
                    strSQL += ", chip_qty = '" & CInt(grid_SlipNoDetail(i, 7)) & "'"
                    If grid_SlipNoDetail(i, 8) = "NO" Then
                        strSQL += ", lot_status = 'Confirm Ready'"
                    Else
                        strSQL += ", lot_status = case when lot_status = 'Confirm Ready' then 'Ready' else lot_status end"
                    End If
                    strSQL += ", confirm = '" & grid_SlipNoDetail(i, 8) & "'"
                    strSQL += ", lot_ww = '" & grid_SlipNoDetail(i, 9) & "'"
                    strSQL += ", step = '" & grid_SlipNoDetail(i, 10) & "'"
                    strSQL += ", issue_date = '" & grid_SlipNoDetail(i, 11) & "'"
                    strSQL += ", reg_date = '" & grid_SlipNoDetail(i, 12) & "'"
                    If grid_SlipNoDetail(i, 13) = String.Empty Then
                        strSQL += ", confirm_date = null"
                    Else
                        strSQL += ", confirm_date = '" & grid_SlipNoDetail(i, 13) & "'"
                    End If
                    strSQL += ", fab_line = '" & grid_SlipNoDetail(i, 14) & "'"
                    strSQL += ", return_type = '" & grid_SlipNoDetail(i, 15) & "'"
                    strSQL += ", lot_option = '" & grid_SlipNoDetail(i, 16) & "'"
                    strSQL += ", note = '" & grid_SlipNoDetail(i, 18) & "'"
                    strSQL += ", writer = '" & loginID & "'"
                    strSQL += ", write_date = '" & write_date & "'"
                    strSQL += ", pfq_doe = '" & grid_SlipNoDetail(i, 17) & "'"
                    strSQL += " where temp_code = '" & grid_SlipNoDetail(i, 19) & "';"
                ElseIf grid_SlipNoDetail(i, 0).ToString = "D" Then
                    '현재행을 삭제 할 경우
                    'strSQL += "delete from PARTS_IN_LIST where IN_CODE = '" & GRID_Parts_List(i, 17) & "';"
                    '현재 테이블외 다른 테이블의 같은 LOT No도 삭제해야하지 않은가?
                    MsgBox("현재 삭제기능은 제공하지 않습니다.", MsgBoxStyle.Information, form_Msgbox_String)
                End If
            Next

            '##### 입고 리스트를 저장한다. 종료 #####

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            th_LoadingWindow.Abort()
            Thread.Sleep(100)
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        th_LoadingWindow.Abort()
        Thread.Sleep(100)
        MsgBox("저장완료.", MsgBoxStyle.Information, form_Msgbox_String)

        BTN_Search_Click(Nothing, Nothing)
        btn_Save.Enabled = False
        btn_Save2.Enabled = False

    End Sub

    Private Sub BTN_Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Search.Click

        thread_LoadingFormStart()

        btn_Save.Enabled = False
        BTN_File_Load1.Enabled = False
        BTN_File_Load2.Enabled = False
        BTN_File_Load3.Enabled = False
        grid_SlipNoDetail.Rows.Count = 1

        grid_SlipNo.Redraw = False
        grid_SlipNo.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call sp_lot_information(0, '" &
                                Format(dtp_Start.Value, "yyyy-MM-dd") & "', '" &
                                Format(DateAdd(DateInterval.Day, 1, dtp_End.Value), "yyyy-MM-dd") & "', '" &
                                "','" &
                                tb_Slip_No.Text & "')"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            'confirm date가 없으면 빈공란으로 해라
            Dim replace_confirm_date As String = String.Empty
            If Not IsDBNull(sqlDR("confirm_date")) Then
                replace_confirm_date = Format(sqlDR("confirm_date"), "yyyy-MM-dd HH:mm:ss")
            End If
            Dim insert_String As String = grid_SlipNo.Rows.Count & vbTab &
                                          sqlDR("slip_no") & vbTab &
                                          Format(sqlDR("first_write_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          replace_confirm_date
            grid_SlipNo.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        grid_SlipNo.AutoSizeCols()
        grid_SlipNo.Redraw = True
        btn_Save.Enabled = False
        btn_Save2.Enabled = False

        th_LoadingWindow.Abort()

    End Sub

    Private Sub CB_Parts_Section_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        e.SuppressKeyPress = True

    End Sub

    Private Sub GRID_Slip_No_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grid_SlipNo.MouseDoubleClick

        If e.Button = Windows.Forms.MouseButtons.Left And grid_SlipNo.MouseRow > 0 Then

            Dim sel_slip_no As String = grid_SlipNo(grid_SlipNo.MouseRow, 1)

            thread_LoadingFormStart()

            grid_SlipNoDetail.Redraw = False
            grid_SlipNoDetail.Rows.Count = 1

            DBConnect()

            Dim strSQL As String = "call sp_lot_information(1, null, null, null, '" & sel_slip_no & "')"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                Dim replace_confirm_date As String = String.Empty
                If Not IsDBNull(sqlDR("CONFIRM_DATE")) Then
                    replace_confirm_date = Format(sqlDR("CONFIRM_DATE"), "yyyy-MM-dd HH:mm:ss")
                End If
                Dim insert_String As String = grid_SlipNoDetail.Rows.Count & vbTab &
                                              sqlDR("part_division") & vbTab &
                                              sqlDR("slip_no") & vbTab &
                                              sqlDR("product") & vbTab &
                                              sqlDR("lot_no") & vbTab &
                                              sqlDR("lot_type") & vbTab &
                                              Format(sqlDR("wafer_qty"), "#,##0") & vbTab &
                                              Format(sqlDR("chip_qty"), "#,##0") & vbTab &
                                              sqlDR("confirm") & vbTab &
                                              sqlDR("lot_ww") & vbTab &
                                              sqlDR("step") & vbTab &
                                              sqlDR("ISSUE_DATE") & vbTab &
                                              Format(sqlDR("reg_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                              replace_confirm_date & vbTab &
                                              sqlDR("fab_line") & vbTab &
                                              sqlDR("return_type") & vbTab &
                                              sqlDR("lot_option") & vbTab &
                                              sqlDR("pfq_doe") & vbTab &
                                              sqlDR("note") & vbTab &
                                              sqlDR("temp_code")
                grid_SlipNoDetail.AddItem(insert_String)
            Loop
            sqlDR.Close()

            DBClose()

            '총입고된 Total 수량을 계산하여 마지막 행에 표시
            Dim total_chip_count As Integer = 0
            Dim total_wafer_count As Integer = 0
            For i = 1 To grid_SlipNoDetail.Rows.Count - 1
                total_chip_count += CInt(grid_SlipNoDetail(i, 7))
            Next

            grid_SlipNoDetail.AddItem(String.Empty & vbTab &
                                    String.Empty & vbTab &
                                    String.Empty & vbTab &
                                    String.Empty & vbTab &
                                    String.Empty & vbTab &
                                    String.Empty & vbTab &
                                    Format(total_wafer_count, "#,##0") & vbTab &
                                    Format(total_chip_count, "#,##0"))

            grid_SlipNoDetail.Rows(grid_SlipNoDetail.Rows.Count - 1).StyleNew.BackColor = Color.AliceBlue

            grid_SlipNoDetail.AutoSizeCols()
            grid_SlipNoDetail.Redraw = True

            BTN_File_Load2.Enabled = True
            BTN_File_Load3.Enabled = True
            btn_Save.Enabled = True
            btn_Save2.Enabled = True

            th_LoadingWindow.Abort()
        End If

    End Sub

    Private Sub GRID_Parts_List_AfterEdit(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs)

        If grid_SlipNoDetail(e.Row, e.Col) = String.Empty Then Exit Sub

        If Not grid_SlipNoDetail(e.Row, 0).ToString = "N" And Not grid_SlipNoDetail(e.Row, 0).ToString = "M" And Not grid_SlipNoDetail(e.Row, 0).ToString = "D" Then

            '수정된 데이터가 서버와 같은지 다른지 확인한 후 실행된다.
            Dim compare_result As Boolean = Data_Compare(grid_SlipNoDetail(e.Row, 17), grid_SlipNoDetail(0, e.Col), grid_SlipNoDetail(e.Row, e.Col))

            If compare_result = False Then
                grid_SlipNoDetail(e.Row, 0) = "M"
                grid_SlipNoDetail.Rows(e.Row).StyleNew.ForeColor = Color.Red
            Else
                grid_SlipNoDetail(e.Row, 0) = e.Row
                grid_SlipNoDetail.Rows(e.Row).StyleNew.ForeColor = Color.Black
            End If

        End If

        grid_SlipNoDetail.AutoSizeCols()

    End Sub

    Private Function Data_Compare(ByVal in_code As String, ByVal col_name As String, ByVal edit_data As String) As Boolean

        Data_Compare = False

        Select Case col_name
            Case "구분"
                col_name = "lot_division"
            Case "Slip No."
                col_name = "slip_no"
            Case "Product Code"
                col_name = "product"
            Case "Lot No."
                col_name = "lot_no"
            Case "Lot Type"
                col_name = "lot_type"
            Case "Wafer Qty"
                col_name = "wafer_qty"
            Case "Chip Qty"
                col_name = "chip_qty"
            Case "Confirm"
                col_name = "confirm"
            Case "주차"
                col_name = "lot_ww"
            Case "Issue Date"
                col_name = "issue_date"
            Case "등록날짜"
                col_name = "reg_date"
            Case "Confirm Date"
                col_name = "confirm_date"
            Case "Fab Line"
                col_name = "fab_line"
            Case "Return Type"
                col_name = "return_type"
            Case "Option"
                col_name = "lot_option"
            Case "비고"
                col_name = "lot_memo"
            Case "PFQ, DOE 물량"
                col_name = "pfq_doe"
        End Select

        DBConnect()

        Dim strSQL As String = "select " & col_name & " from PARTS_IN_LIST where IN_CODE = '" & in_code & "'"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            If sqlDR(col_name).ToString = edit_data Then
                Data_Compare = True
            End If
        Loop
        sqlDR.Close()

        DBClose()

        Return Data_Compare

    End Function

    Private Sub GRID_Parts_List_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grid_SlipNoDetail.KeyDown

        If grid_SlipNoDetail.Row < 0 Then Exit Sub
        If grid_SlipNoDetail(grid_SlipNoDetail.Row, 0).ToString = "D" Then
            e.SuppressKeyPress = True
        End If

    End Sub

    Private Sub GRID_Parts_List_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grid_SlipNoDetail.MouseClick

        If e.Button = Windows.Forms.MouseButtons.Right And grid_SlipNoDetail.MouseRow > -1 Then
            '클릭한 행을 선택행으로
            grid_SlipNoDetail.Row = grid_SlipNoDetail.MouseRow
            '현재행이 삭제 지정된 행인지 아닌지에 따라 '삭제취소'버튼의 활성화 여부 결정
            If grid_SlipNoDetail(grid_SlipNoDetail.Row, 0).ToString = "D" Then
                btn_Delete_Cancel.Enabled = True
            Else
                btn_Delete_Cancel.Enabled = False
            End If

            cms_Menu.Show(grid_SlipNoDetail, New Point(e.X, e.Y))
        End If

    End Sub

    Private Sub Grid_New_Index()

        grid_SlipNoDetail.Redraw = False

        For i = 1 To grid_SlipNoDetail.Rows.Count - 1
            If IsNumeric(grid_SlipNoDetail(i, 0)) Then
                grid_SlipNoDetail(i, 0) = i
                grid_SlipNoDetail.Rows(i).StyleNew.ForeColor = Color.Black
            Else
                If grid_SlipNoDetail(i, 0).ToString = "N" Then
                    grid_SlipNoDetail.Rows(i).StyleNew.ForeColor = Color.Blue
                ElseIf grid_SlipNoDetail(i, 0).ToString = "D" Then
                    grid_SlipNoDetail.Rows(i).StyleNew.ForeColor = Color.Gray
                End If
            End If
        Next

        grid_SlipNoDetail.AutoSizeCols()

        grid_SlipNoDetail.Redraw = True

    End Sub

    Public Sub Row_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_RowAdd.Click

        Dim c1f As C1FlexGrid = grid_SlipNoDetail

        Dim temp_In_Code As String = "MCIN" & Format(Now, "yyMMddHHmmss01")

        For i = 1 To c1f.Rows.Count - 1
            If c1f(i, 18) = temp_In_Code Then '같은 입고코드가 존재한다면.. 변경
                temp_In_Code = Replace(temp_In_Code, "MCIN", String.Empty)
                temp_In_Code = "MCIN" & CDbl(temp_In_Code) + 1
            End If
        Next

        c1f.AddItem("N" & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            temp_In_Code, c1f.Row + 1)
        Grid_New_Index()

    End Sub

    Private Sub Row_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_RowDelete.Click

        If grid_SlipNoDetail(grid_SlipNoDetail.Row, 0).ToString = "N" Then
            grid_SlipNoDetail.RemoveItem(grid_SlipNoDetail.Row)
        Else
            grid_SlipNoDetail(grid_SlipNoDetail.Row, 0) = "D"
        End If

        Grid_New_Index()

    End Sub

    Private Sub Delete_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Delete_Cancel.Click

        If grid_SlipNoDetail(grid_SlipNoDetail.Row, 0).ToString = "D" Then
            grid_SlipNoDetail(grid_SlipNoDetail.Row, 0) = grid_SlipNoDetail.Row
        End If

        Grid_New_Index()

    End Sub

    Private Sub Form_CLose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub Multi_Row_Insert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Multi_Row_Insert.Click

        If Not Multi_Row.Visible Then Multi_Row.Show()
        Multi_Row.Focus()

    End Sub

    Private Sub grid_SlipNoDetail_AfterEdit(sender As Object, e As RowColEventArgs) Handles grid_SlipNoDetail.AfterEdit

        If IsNumeric(grid_SlipNoDetail(e.Row, 0)) Then
            grid_SlipNoDetail(e.Row, 0) = "M"
            grid_SlipNoDetail.Rows(e.Row).StyleNew.ForeColor = Color.Red
        End If

        grid_SlipNoDetail.Redraw = False
        grid_SlipNoDetail.AutoSizeCols()
        grid_SlipNoDetail.Redraw = True

    End Sub

    Private Sub grid_SlipNo_MouseMove(sender As Object, e As MouseEventArgs) Handles grid_SlipNo.MouseMove,
                                                                                     grid_SlipNoDetail.MouseMove

        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then
            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub

    Private Sub grid_SlipNoDetail_LostFocus(sender As Object, e As EventArgs) Handles grid_SlipNoDetail.LostFocus,
                                                                                      grid_SlipNo.LostFocus

        frm_Main.lb_Status.Text = String.Empty

    End Sub
End Class
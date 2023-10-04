'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-08-31
'##### 수정일자 : 2023-08-31
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 

'############################################################################################################
'############################################################################################################

Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient
Imports System.ComponentModel
Imports System.Threading

Public Class frm_Monthly_Production_Data

    Dim form_Msgbox_String As String = "OMS 생산월보 저장"
    Dim runProcess As Thread
    Dim open_Cancel As Boolean
    Dim openFilePath As String
    Dim first_open As Boolean = False
    Dim first_run As Boolean = True

    Private Sub frm_Monthly_Production_Data_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        GRID_SETTING()

        DateTimePicker1.Value = Format(DateAdd(DateInterval.Day, -1, Now), "yyyy-MM-dd")

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub GRID_SETTING()

        With grid_SlipNoDetail
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
            grid_SlipNoDetail(0, 0) = "No"
            grid_SlipNoDetail(0, 1) = "일자"
            grid_SlipNoDetail(0, 2) = "Product"
            grid_SlipNoDetail(0, 3) = "Fab Line"
            grid_SlipNoDetail(0, 4) = "Lot Type"
            grid_SlipNoDetail(0, 5) = "Return Type"
            grid_SlipNoDetail(0, 6) = "Chip Qty"
            .Cols(6).Format = "#,##0"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .AutoSizeCols()
            .ShowCursor = True
            .SubtotalPosition = SubtotalPositionEnum.BelowData
        End With

        'set up styles
        Dim s As CellStyle = grid_SlipNoDetail.Styles(CellStyleEnum.Subtotal0)
        s.BackColor = Color.DarkBlue
        s.ForeColor = Color.White
        s.Font = New Font(grid_SlipNoDetail.Font, FontStyle.Bold)

        s = grid_SlipNoDetail.Styles(CellStyleEnum.Subtotal1)
        s.BackColor = Color.LightBlue
        s.ForeColor = Color.Black
        s = grid_SlipNoDetail.Styles(CellStyleEnum.Subtotal2)
        s.BackColor = Color.DarkRed
        s.ForeColor = Color.White

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        open_Cancel = True

    End Sub

    Private Sub btn_File_Load_Click(sender As Object, e As EventArgs) Handles btn_File_Load.Click

        grid_SlipNoDetail.Rows.Count = 1

        Dim OpenFileDialog1 As New OpenFileDialog

        OpenFileDialog1.Filter = "Excel (*.xls, *.xlsx)|*.xls;*.xlsx"
        OpenFileDialog1.InitialDirectory = "D:\1. Workplace\5. 고객사 자료\18. 삼성전자(Repair)\##1. OMS Upload\3. OMS Excel Download\3. 생산월보"
        OpenFileDialog1.ShowDialog()

        If OpenFileDialog1.FileName = String.Empty Then Exit Sub

        OpenFileDialog1.Dispose()

        File_Open_Start()
        openFilePath = OpenFileDialog1.FileName

    End Sub

    Private Sub File_Open_Start()

        runProcess = New Thread(AddressOf File_Open)
        runProcess.IsBackground = True
        runProcess.SetApartmentState(ApartmentState.STA) 'OpenFileDialog를 사용하기위해선 STA로 해야되던데...
        runProcess.Start()

    End Sub

    Private Sub File_Open()

        Invoke(SetGridRedraw, False)
        open_Cancel = False

        Invoke(ButtonEnabled, btn_File_Load, False)

        Dim excel_App As Object

        excel_App = CreateObject("Excel.Application")
        excel_App.WorkBooks.Open(openFilePath)

        excel_App.Visible = False

        Dim sheet_no As Integer = 1
        Dim date_OMS As String = Format(DateTimePicker1.Value, "yyyy-MM-dd")

        With excel_App.ActiveWorkbook.Sheets(sheet_no)
            For i = CInt(tb_StartRow.Text) To .UsedRange.Rows.Count - 1
                If open_Cancel = True Then '불러오기 중단을 요청한 경우
                    Invoke(MsgboxRun, "데이터 불러오기를 중단 하였습니다.")
                    Invoke(ButtonEnabled, btn_File_Load, True)
                    excel_App.WorkBooks(1).Close()
                    excel_App.Quit()
                    excel_App = Nothing
                    runProcess.Abort()
                End If

                Try
                    Dim replace_ReturnType As String = .Cells(i, CInt(tb_ReturnType.Text)).Value
                    If replace_ReturnType.Equals("**") Then
                        replace_ReturnType = String.Empty
                    End If

                    Dim oms_Product As String = .Cells(i, CInt(tb_Product.Text)).Value
                    Dim oms_FabLine As String = .Cells(i, CInt(tb_FabLine.Text)).Value
                    Dim oms_LotType As String = .Cells(i, CInt(tb_LotType.Text)).Value
                    If oms_LotType.Equals("P") Then oms_LotType = "PP"
                    Dim oms_Loss As Integer = Format(CInt(.Cells(i, CInt(tb_Loss.Text)).Value), "#,##0")
                    'product 칸에 total이 오면 불러오지 않는다.
                    If Not oms_Product Like "MODULE*" Then
                        Invoke(SetPartsList, "N" & vbTab &
                                         date_OMS & vbTab &
                                         oms_Product & vbTab &
                                         oms_FabLine & vbTab &
                                         oms_LotType & vbTab &
                                         replace_ReturnType & vbTab &
                                         oms_Loss)
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
                    excel_App.WorkBooks(1).Close()
                    excel_App.Quit()
                    excel_App = Nothing

                    Invoke(ButtonEnabled, btn_File_Load, True)
                    runProcess.Abort()
                End Try

            Next
        End With

        excel_App.WorkBooks(1).Close()
        excel_App.Quit()
        excel_App = Nothing

        Invoke(ButtonEnabled, btn_File_Load, True)
        Invoke(SetGridRedraw, True)

        Invoke(SetNewLine)

        Invoke(SetGridSubTotal)

        Invoke(MsgboxRun, "완료 하였습니다." & vbCrLf & "데이터를 확인하여 주십시오.")

        runProcess.Abort()

    End Sub

    Private Delegate Sub GridRSubTotalDelegate()
    Private SetGridSubTotal = New GridRSubTotalDelegate(AddressOf UpdateTotals)

    Private Sub UpdateTotals()
        'clear existing totals
        grid_SlipNoDetail.Subtotal(AggregateEnum.Clear)

        'calculate subtotals (three levels, totals on every column)
        'Dim c As Integer
        'For c = 3 To C1FlexGrid1.Cols.Count - 1
        'grid_PartHistory.Subtotal(AggregateEnum.Sum, 0, 1, 4, "Grand Total")
        grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, -1, -1, 6, "Grand Total")
        'grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 0, 1, 6, "Grand Total")
        'grid_PartHistory.Subtotal(AggregateEnum.Sum, 1, 3, 4, "Total for {0}")
        'grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 1, 3, 5, "Total for {0}")
        'grid_SlipNoDetail.Subtotal(AggregateEnum.Sum, 1, 3, 6, "Total for {0}")
        'C1FlexGrid1.Subtotal(AggregateEnum.Sum, 2, 2, 5, "Total for {0}")
        'Next

        'done, autosize columns to finish
        grid_SlipNoDetail.AutoSizeCols()

    End Sub

    Private Delegate Sub GridRedrawDelegate(ByVal status As Boolean)
    Private SetGridRedraw = New GridRedrawDelegate(AddressOf GridRedraw)

    Private Sub GridRedraw(ByVal status As Boolean)

        grid_SlipNoDetail.Redraw = status

    End Sub

    Private Delegate Sub ButtonEnabledDelegate(ByVal buttonName As Button, ByVal status As Boolean)
    Private ButtonEnabled = New ButtonEnabledDelegate(AddressOf ButtonStatus)

    Private Sub ButtonStatus(ByVal ButtonName As Button, ByVal Status As Boolean)

        ButtonName.Enabled = Status

    End Sub

    Private Delegate Sub MsgBoxDelegate(ByVal ShowText As String)
    Private MsgboxRun = New MsgBoxDelegate(AddressOf MsgShow)

    Private Sub MsgShow(ByVal ShowText As String)

        MsgBox(ShowText, MsgBoxStyle.Information, form_Msgbox_String)

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

    Private Delegate Sub SetPartsListDelegate(ByVal addText As String)
    Private SetPartsList As SetPartsListDelegate = New SetPartsListDelegate(AddressOf GridROW)

    Private Sub GridROW(ByVal AddText As String)

        grid_SlipNoDetail.AddItem(AddText)
        grid_SlipNoDetail.AutoSizeCols()
        grid_SlipNoDetail.AutoSizeRows(1, 0, grid_SlipNoDetail.Rows.Count - 1, grid_SlipNoDetail.Cols.Count - 1, 0, AutoSizeFlags.None)
        grid_SlipNoDetail.Rows(0).Height = 30

    End Sub

    Private Sub btn_Save2_Click(sender As Object, e As EventArgs) Handles btn_Save2.Click

        If MsgBox("저장 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Saving...")

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = "delete from oms_production_monthly_report"
        strSQL += " where data_date = '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "';"

        sqlTran = dbConnection1.BeginTransaction

        Try

            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            For i = 1 To grid_SlipNoDetail.Rows.Count - 1
                If Not IsNothing(grid_SlipNoDetail(i, 0)) Then
                    If grid_SlipNoDetail(i, 0).ToString = "N" Then
                        '현재행이 신규 입력행일 경우
                        strSQL += "insert into oms_production_monthly_report(data_date, product"
                        strSQL += ", fab_line, lot_type, return_type, chip_qty) values("
                        strSQL += "'" & grid_SlipNoDetail(i, 1) & "'"
                        strSQL += ",'" & grid_SlipNoDetail(i, 2) & "'"
                        strSQL += ",'" & grid_SlipNoDetail(i, 3) & "'"
                        strSQL += ",'" & grid_SlipNoDetail(i, 4) & "'"
                        strSQL += ",'" & grid_SlipNoDetail(i, 5) & "'"
                        strSQL += ",'" & grid_SlipNoDetail(i, 6) & "'"
                        strSQL += ");"
                    End If
                End If
            Next

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

        thread_LoadingFormEnd
        Thread.Sleep(100)

        MsgBox("저장완료.", MsgBoxStyle.Information, form_Msgbox_String)

        btn_Save2.Enabled = True

        DateTimePicker1_ValueChanged(Nothing, Nothing)

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        thread_LoadingFormStart()

        grid_SlipNoDetail.Rows.Count = 1
        grid_SlipNoDetail.Redraw = False

        DBConnect()

        Dim strSQL As String = "select data_date, product, fab_line, lot_type, return_type, chip_qty from oms_production_monthly_report"
        strSQL += " where data_date = '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "';"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insertString As String = grid_SlipNoDetail.Rows.Count & vbTab &
                sqlDR("data_date") & vbTab &
                sqlDR("product") & vbTab &
                sqlDR("fab_line") & vbTab &
                sqlDR("lot_type") & vbTab &
                sqlDR("return_type") & vbTab &
                Format(sqlDR("chip_qty"), "#,##0")
            grid_SlipNoDetail.AddItem(insertString)
        Loop
        sqlDR.Close()

        DBClose()

        UpdateTotals()

        grid_SlipNoDetail.AutoSizeCols()
        grid_SlipNoDetail.Redraw = True

        thread_LoadingFormEnd

    End Sub

    Private Sub grid_SlipNoDetail_MouseClick(sender As Object, e As MouseEventArgs) Handles grid_SlipNoDetail.MouseClick

        If grid_SlipNoDetail.MouseRow > 0 And e.Button = MouseButtons.Right Then
            cms_Menu.Show(grid_SlipNoDetail, New Point(e.X, e.Y))
        End If

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
        Else
            frm_Main.lb_Status.Text = String.Empty
        End If

    End Sub
End Class
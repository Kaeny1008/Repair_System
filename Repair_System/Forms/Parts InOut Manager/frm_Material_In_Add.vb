'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-07-20
'##### 수정일자 : 2023-07-20
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : Module, Component 외 입고등록을 한다.

'############################################################################################################
'############################################################################################################

Imports System.Threading
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_Material_In_Add

    Dim form_Msgbox_String As String = "자재 입고 등록"
    Dim runProcess As Thread

    Private Sub PARTS_IN2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        dtp_Start.Value = Format(Now, "yyyy-MM-01")
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
            .Cols.Count = 3
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_SlipNo(0, 0) = "No"
            grid_SlipNo(0, 1) = "Slip No."
            grid_SlipNo(0, 2) = "입고일자(시간)"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

        With grid_Parts_List
            .AllowEditing = True
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 8
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_Parts_List(0, 0) = "No"
            grid_Parts_List(0, 1) = "구분"
            grid_Parts_List(0, 2) = "Slip No."
            grid_Parts_List(0, 3) = "Part No."
            grid_Parts_List(0, 4) = "Lot No."
            grid_Parts_List(0, 5) = "Qty"
            grid_Parts_List(0, 6) = "비고"
            grid_Parts_List(0, 7) = "temp_code"
            .Cols(5).DataType = GetType(Integer)
            .Cols(5).Format = "#,##0"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 2).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .Cols(.Cols.Count - 1).Visible = False
            .ExtendLastCol = True
            .ShowCursor = True
            .AutoSizeCols()
        End With

    End Sub

    Private Sub GRID_Parts_List_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grid_Parts_List.MouseClick

        Dim c1f As C1FlexGrid = grid_Parts_List

        If e.Button = Windows.Forms.MouseButtons.Right And c1f.MouseRow > -1 Then
            c1f.Row = c1f.MouseRow
            If c1f(c1f.Row, 0).ToString = "D" Then
                Row_Delete.Enabled = False
                Delete_Cancel.Enabled = True
            Else
                Row_Delete.Enabled = True
                Delete_Cancel.Enabled = False
            End If
            cms_GridMenu.Show(c1f, New Point(e.X, e.Y))
        End If

    End Sub

    Private Sub Grid_New_Index()

        Dim c1f As C1FlexGrid = grid_Parts_List

        c1f.Redraw = False

        For i = 1 To c1f.Rows.Count - 1
            If IsNumeric(c1f(i, 0)) Then
                c1f(i, 0) = i
                c1f.Rows(i).StyleNew.ForeColor = Color.Black
            Else
                If c1f(i, 0).ToString = "N" Then
                    c1f.Rows(i).StyleNew.ForeColor = Color.Blue
                ElseIf c1f(i, 0).ToString = "D" Then
                    c1f.Rows(i).StyleNew.ForeColor = Color.Gray
                End If
            End If
        Next

        c1f.AutoSizeCols()

        c1f.Redraw = True

    End Sub

    Private Sub Row_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Row_ADD.Click

        Dim c1f As C1FlexGrid = grid_Parts_List

        Dim temp_In_Code As String = "MAIN" & Format(Now, "yyMMddHHmmss01")

        For i = 1 To c1f.Rows.Count - 1
            If c1f(i, 7) = temp_In_Code Then '같은 입고코드가 존재한다면.. 변경
                temp_In_Code = Replace(temp_In_Code, "MAIN", String.Empty)
                temp_In_Code = "MAIN" & CDbl(temp_In_Code) + 1
            End If
        Next

        c1f.AddItem("N" & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            String.Empty & vbTab &
            temp_In_Code, c1f.Row + 1)
        Grid_New_Index()

    End Sub

    Private Sub Row_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Row_Delete.Click

        Dim c1f As C1FlexGrid = grid_Parts_List

        If c1f(c1f.Row, 0).ToString = "N" Then
            c1f.RemoveItem(c1f.Row)
        Else
            c1f(c1f.Row, 0) = "D"
        End If

        Grid_New_Index()

    End Sub

    Private Sub Delete_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Delete_Cancel.Click

        Dim c1f As C1FlexGrid = grid_Parts_List

        If c1f(c1f.Row, 0).ToString = "D" Then
            c1f(c1f.Row, 0) = c1f.Row
        End If

        Grid_New_Index()

    End Sub

    Private Sub btn_NewParts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_NewParts.Click

        grid_Parts_List.Rows.Count = 1
        btn_File_Load.Enabled = True
        BTN_Save.Enabled = True
        BTN_Save2.Enabled = True

    End Sub

    Private Sub btn_File_Load_Click(sender As Object, e As EventArgs) Handles btn_File_Load.Click

        grid_Parts_List.Rows.Count = 1

        If Not frm_P_First_Popup.Visible Then frm_P_First_Popup.Show()
        frm_P_First_Popup.Focus()

    End Sub

    Public slip_no, part_no, lot_no, qty, start_row As Integer
    Public open_File_Path As String
    Dim open_Cancel As Boolean

    Public Sub File_Open_Start()

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
                    Invoke(ButtonEnabled, btn_File_Load, True)
                    excel_App.WorkBooks(1).Close()
                    excel_App.Quit()
                    excel_App = Nothing
                    runProcess.Abort()
                End If

                Invoke(SetPGStatus, "데이터를 불러오고 있습니다.     " & i - start_row + 1 & " / " & .UsedRange.Rows.Count - start_row & " 행") '현재 진행상태를 표시

                Dim r_slip_no As String = .Cells(i, slip_no).Value
                Dim r_part_no As String = .Cells(i, part_no).Value
                Dim r_lot_no As String = .Cells(i, lot_no).Value
                Dim r_qty As String = .Cells(i, qty).Value
                Dim item_division As String = String.Empty

                If Invoke(divisionFind, "pmic_partno", r_part_no) > 0 Then
                    item_division = "PMIC"
                ElseIf Invoke(divisionFind, "rcd_partno", r_part_no) > 0 Then
                    item_division = "RCD"
                End If

                write_code += 1

                Try
                    Invoke(SetPartsList, "N" & vbTab &
                                         item_division & vbTab &
                                         r_slip_no & vbTab &
                                         r_part_no & vbTab &
                                         r_lot_no & vbTab &
                                         Format(CInt(r_qty), "#,##0") & vbTab &
                                         String.Empty & vbTab &
                                         "MAIN" & write_code)
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

        Invoke(SetPGStatus, "성공적으로 데이터를 불러왔습니다.") '현재 진행상태를 표시

        Invoke(ButtonEnabled, btn_File_Load, True)
        Invoke(SetGridRedraw, True)

        Invoke(SetNewLine)

        Invoke(MsgboxRun, "완료 하였습니다." & vbCrLf & "데이터를 확인하여 주십시오.")

        runProcess.Abort()

    End Sub

    Private Delegate Function partDivisionFindDelegate(ByVal section As String, ByVal part_no As String) As Integer
    Private divisionFind = New partDivisionFindDelegate(AddressOf partDivisionFind)

    Private Function partDivisionFind(ByVal section As String, ByVal part_no As String) As Integer

        DBConnect()

        Dim strSQL As String = "call sp_incoming_inspect(null, '" & section & "', '" & part_no & "', 3)"
        Dim total_count As Integer = 0

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            total_count += 1
        Loop
        sqlDR.Close()

        DBClose()

        Return total_count

    End Function

    Private Delegate Sub GridRedrawDelegate(ByVal status As Boolean)
    Private SetGridRedraw = New GridRedrawDelegate(AddressOf GridRedraw)

    Private Sub GridRedraw(ByVal status As Boolean)

        grid_Parts_List.Redraw = status

    End Sub

    Private Delegate Sub ButtonEnabledDelegate(ByVal buttonName As Button, ByVal status As Boolean)
    Private ButtonEnabled = New ButtonEnabledDelegate(AddressOf ButtonStatus)

    Private Sub ButtonStatus(ByVal ButtonName As Button, ByVal Status As Boolean)

        ButtonName.Enabled = Status

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

    Private Delegate Sub SetPartsListDelegate(ByVal addText As String)
    Private SetPartsList As SetPartsListDelegate = New SetPartsListDelegate(AddressOf GridROW)

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub btn_LoadStop_Click(sender As Object, e As EventArgs) Handles btn_LoadStop.Click

        open_Cancel = True

    End Sub

    Private Sub BTN_Save_Click(sender As Object, e As EventArgs) Handles BTN_Save.Click, BTN_Save2.Click

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

            For i = 1 To grid_Parts_List.Rows.Count - 1
                If grid_Parts_List(i, 0).ToString = "N" Then
                    '현재행이 신규 입력행일 경우
                    strSQL += "insert into basic_material_information(part_division, slip_no, part_no, lot_no, material_qty, material_note"
                    strSQL += ", temp_code, first_write_date, writer, write_date) values"
                    strSQL += "('" & grid_Parts_List(i, 1) & "'" '구분
                    strSQL += ",'" & grid_Parts_List(i, 2) & "'" 'Slip No
                    strSQL += ",'" & grid_Parts_List(i, 3) & "'" 'part no
                    strSQL += ",'" & grid_Parts_List(i, 4) & "'" 'Lot No.
                    strSQL += ",'" & CInt(grid_Parts_List(i, 5)) & "'" 'Qty
                    strSQL += ",'" & grid_Parts_List(i, 6) & "'" '비고
                    strSQL += ",'" & grid_Parts_List(i, 7) & "'" 'temp code
                    strSQL += ",'" & write_date & "'"
                    strSQL += ",'" & loginID & "'"
                    strSQL += ",'" & write_date & "');"
                ElseIf grid_Parts_List(i, 0).ToString = "M" Then
                    '현재행이 수정된 경우
                    strSQL += "update basic_material_information set part_division = '" & grid_Parts_List(i, 1) & "'"
                    strSQL += ", slip_no = '" & grid_Parts_List(i, 2) & "'"
                    strSQL += ", part_no = '" & grid_Parts_List(i, 3) & "'"
                    strSQL += ", lot_no = '" & grid_Parts_List(i, 4) & "'"
                    strSQL += ", material_qty = '" & CInt(grid_Parts_List(i, 5)) & "'"
                    strSQL += ", material_note = '" & grid_Parts_List(i, 6) & "'"
                    strSQL += " where temp_code = '" & grid_Parts_List(i, 7) & "';"
                ElseIf grid_Parts_List(i, 0).ToString = "D" Then
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
            thread_LoadingFormEnd
            Thread.Sleep(100)
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        thread_LoadingFormEnd
        Thread.Sleep(100)
        MsgBox("저장완료.", MsgBoxStyle.Information, form_Msgbox_String)

        BTN_Search_Click(Nothing, Nothing)
        BTN_Save.Enabled = False
        BTN_Save2.Enabled = False

    End Sub

    Private Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        thread_LoadingFormStart()

        BTN_Save.Enabled = False
        btn_File_Load.Enabled = False
        grid_Parts_List.Rows.Count = 1

        grid_SlipNo.Redraw = False
        grid_SlipNo.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call sp_material_information(0, '" &
                                Format(dtp_Start.Value, "yyyy-MM-dd") & "', '" &
                                Format(DateAdd(DateInterval.Day, 1, dtp_End.Value), "yyyy-MM-dd") & "', '" &
                                "','" &
                                TB_Slip_No.Text & "')"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_SlipNo.Rows.Count & vbTab &
                                          sqlDR("slip_no") & vbTab &
                                          Format(sqlDR("first_write_date"), "yyyy-MM-dd HH:mm:ss")
            grid_SlipNo.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        grid_SlipNo.AutoSizeCols()
        grid_SlipNo.Redraw = True
        BTN_Save.Enabled = False
        BTN_Save2.Enabled = False

        thread_LoadingFormEnd

    End Sub

    Private Sub GridROW(ByVal AddText As String)

        grid_Parts_List.AddItem(AddText)
        grid_Parts_List.AutoSizeCols()
        grid_Parts_List.AutoSizeRows()
        grid_Parts_List.Rows(0).Height = 30

    End Sub

    Private Delegate Sub NewLineDelegate()
    Private SetNewLine = New NewLineDelegate(AddressOf GRID_NewLine)

    Private Sub GRID_NewLine()

        For i = 1 To grid_Parts_List.Rows.Count - 1
            If grid_Parts_List(i, 0).ToString = "N" Then
                grid_Parts_List.Rows(i).StyleNew.ForeColor = Color.Blue
            End If
        Next

    End Sub

    Private Sub grid_Parts_List_AfterEdit(sender As Object, e As RowColEventArgs) Handles grid_Parts_List.AfterEdit

        If IsNumeric(grid_Parts_List(e.Row, 0)) Then
            grid_Parts_List(e.Row, 0) = "M"
            grid_Parts_List.Rows(e.Row).StyleNew.ForeColor = Color.Red
        End If

        If grid_Parts_List(e.Row, 0) = "N" Then
            If partDivisionFind("pmic_partno", grid_Parts_List(e.Row, 3)) > 0 Then
                grid_Parts_List(e.Row, 1) = "PMIC"
            ElseIf partDivisionFind("rcd_partno", grid_Parts_List(e.Row, 3)) > 0 Then
                grid_Parts_List(e.Row, 1) = "RCD"
            Else
                grid_Parts_List(e.Row, 1) = String.Empty
                MsgBox("Part No 정보를 찾을 수 없습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            End If
        End If

            grid_Parts_List.Redraw = False
        grid_Parts_List.AutoSizeCols()
        grid_Parts_List.Redraw = True

    End Sub

    Private Sub grid_SlipNo_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles grid_SlipNo.MouseDoubleClick

        If e.Button = Windows.Forms.MouseButtons.Left And grid_SlipNo.MouseRow > 0 Then

            Dim sel_slip_no As String = grid_SlipNo(grid_SlipNo.MouseRow, 1)

            grid_Parts_List.Redraw = False
            grid_Parts_List.Rows.Count = 1

            DBConnect()

            Dim strSQL As String = "call sp_material_information(1, null, null, null, '" & sel_slip_no & "')"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                Dim insert_String As String = grid_Parts_List.Rows.Count & vbTab &
                                              sqlDR("part_division") & vbTab &
                                              sqlDR("slip_no") & vbTab &
                                              sqlDR("part_no") & vbTab &
                                              sqlDR("lot_no") & vbTab &
                                              Format(sqlDR("material_qty"), "#,##0") & vbTab &
                                              sqlDR("material_note") & vbTab &
                                              sqlDR("temp_code")
                grid_Parts_List.AddItem(insert_String)
            Loop
            sqlDR.Close()

            DBClose()

            '총입고된 Total 수량을 계산하여 마지막 행에 표시
            Dim total_chip_count As Integer = 0
            For i = 1 To grid_Parts_List.Rows.Count - 1
                total_chip_count += CInt(grid_Parts_List(i, 5))
            Next

            grid_Parts_List.AddItem(String.Empty & vbTab &
                                    String.Empty & vbTab &
                                    String.Empty & vbTab &
                                    String.Empty & vbTab &
                                    String.Empty & vbTab &
                                    Format(total_chip_count, "#,##0"))

            grid_Parts_List.Rows(grid_Parts_List.Rows.Count - 1).StyleNew.BackColor = Color.AliceBlue

            grid_Parts_List.AutoSizeCols()
            grid_Parts_List.Redraw = True

            BTN_Save.Enabled = True
            BTN_Save2.Enabled = True
        End If

    End Sub
End Class
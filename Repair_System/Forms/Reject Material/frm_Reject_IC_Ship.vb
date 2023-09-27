'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-08-29
'##### 수정일자 : 2023-08-29
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 

'############################################################################################################
'############################################################################################################

Imports System.Threading
Imports C1.Win.C1FlexGrid
Imports CrystalDecisions.CrystalReports.Engine
Imports MySql.Data.MySqlClient

Public Class frm_Reject_IC_Ship

    Dim form_Msgbox_String As String = "폐기자재 반납"

    Private Sub frm_Reject_IC_Ship_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        grid_Setting()

        DateTimePicker1.Value = Format(Now, "yyyy-MM-01 00:00:00")
        DateTimePicker2.Value = Format(Now, "yyyy-MM-dd 23:59:59")
        date_Reject_Start.Value = Format(Now, "yyyy-MM-01 00:00:00")
        date_Reject_End.Value = Format(Now, "yyyy-MM-dd 23:59:59")

    End Sub

    Private Sub grid_Setting()

        With grid_RejectList
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
            grid_RejectList(0, 0) = "No"
            grid_RejectList(0, 1) = "반납 일자"
            grid_RejectList(0, 2) = "반납번호"
            grid_RejectList(0, 3) = "Lot 수"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

        With grid_RejectList_Detail
            .AllowEditing = False
            .AllowFiltering = False
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 10
            .Cols.Fixed = 1
            .Rows.Count = 2
            .Rows.Fixed = 2
            .AutoClipboard = True
            .AllowFreezing = AllowFreezingEnum.Both
            .AllowMerging = AllowMergingEnum.FixedOnly
            .AllowMergingFixed = AllowMergingEnum.FixedOnly
            For i = 0 To .Cols.Count - 1
                .Cols(i).AllowMerging = True
            Next
            .Rows(0).AllowMerging = True
            .Rows(1).AllowMerging = True
            Dim Mrng As CellRange = .GetCellRange(0, 0, 1, 0)
            Mrng.Data = "No"
            Mrng = .GetCellRange(0, 1, 1, 1)
            Mrng.Data = "YJ No"
            Mrng = .GetCellRange(0, 2, 1, 4)
            Mrng.Data = "Module"
            grid_RejectList_Detail(1, 2) = "Product Code"
            grid_RejectList_Detail(1, 3) = "Lot No."
            grid_RejectList_Detail(1, 4) = "수량(EA)"
            Mrng = .GetCellRange(0, 5, 1, 8)
            Mrng.Data = "반납정보"
            grid_RejectList_Detail(1, 5) = "PMIC"
            grid_RejectList_Detail(1, 6) = "반납수량"
            grid_RejectList_Detail(1, 7) = "RCD"
            grid_RejectList_Detail(1, 8) = "반납수량"
            Mrng = .GetCellRange(0, 9, 1, 9)
            Mrng.Data = "비고"
            '.Styles("Normal").WordWrap = True
            '.Cols(11).StyleNew.WordWrap = True

            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 1).TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ExtendLastCol = True
            .ShowCursor = True
        End With

        With grid_PartList
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
            grid_PartList(0, 0) = "No"
            grid_PartList(0, 1) = "구분"
            grid_PartList(0, 2) = "Part No."
            grid_PartList(0, 3) = "Lot 수(Lot)"
            grid_PartList(0, 4) = "수량(EA)"
            grid_PartList(0, 5) = "무게(Kg)"
            grid_PartList(0, 6) = "비고"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
            .ExtendLastCol = True
        End With

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub btn_New_Reject_Click(sender As Object, e As EventArgs) Handles btn_New_Reject.Click

        TextBox2.Text = "여기는 누락된 데이터가 표시되는 곳입니다."
        RS_Code_Maker()
        btn_Save.Enabled = True

        grid_RejectList_Detail.Rows.Count = 2
        grid_PartList.Rows.Count = 1

        tb_Min_No.Enabled = True
        tb_Max_No.Enabled = True

        date_Reject_Start.Enabled = True
        date_Reject_End.Enabled = True

        tb_Min_No.SelectAll()
        tb_Min_No.Focus()
        btn_PMIC_Print.Enabled = False

    End Sub

    Private Sub RS_Code_Maker()

        Dim rn_code As String = String.Empty

        DBConnect()

        Try
            Dim strSQL As String = "Call sp_reject_material_ship(0" &
                        ", '" & Format(Now, "yyyy-MM-dd") & "'" &
                        ", null" &
                        ", null" &
                        ", null)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                If Not IsDBNull(sqlDR("ship_no")) Then rn_code = "RS" & Format(Now, "yyMMdd") & sqlDR("ship_no")
            Loop
            sqlDR.Close()
        Catch ex As Exception
            MsgBox("반납번호 생성중 에러 발생 : " & vbCrLf & ex.Message,
                   MsgBoxStyle.Critical,
                   form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        If rn_code = String.Empty Then rn_code = "RS" & Format(Now, "yyMMdd") & "01"

        TextBox1.Text = rn_code

    End Sub

    Private Sub tb_Min_No_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_Min_No.KeyDown

        Select Case e.KeyCode
            Case 8, 96 To 105, 48 To 57
                rb_Area.Checked = True
            Case 13
                rb_Area.Checked = True
                If Not Trim(tb_Min_No.Text) = String.Empty And
                    Trim(tb_Max_No.Text) = String.Empty Then
                    tb_Max_No.SelectAll()
                    tb_Max_No.Focus()
                End If
                If Not Trim(tb_Min_No.Text) = String.Empty And Not _
                    Trim(tb_Max_No.Text) = String.Empty Then
                    btn_InputCheck_Click(Nothing, Nothing)
                End If
            Case Else
                e.SuppressKeyPress = True
        End Select

    End Sub

    Private Sub tb_Max_No_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_Max_No.KeyDown

        Select Case e.KeyCode
            Case 8, 96 To 105, 48 To 57
                rb_Area.Checked = True
            Case 13
                rb_Area.Checked = True
                If Not Trim(tb_Max_No.Text) = String.Empty And
                    Trim(tb_Min_No.Text) = String.Empty Then
                    tb_Min_No.Focus()
                End If
                If Not Trim(tb_Min_No.Text) = String.Empty And Not _
                    Trim(tb_Max_No.Text) = String.Empty Then
                    btn_InputCheck_Click(Nothing, Nothing)
                End If
            Case Else
                e.SuppressKeyPress = True
        End Select

    End Sub

    Private Sub btn_InputCheck_Click(sender As Object, e As EventArgs) Handles btn_InputCheck.Click

        thread_LoadingFormStart()

        If rb_Area.Checked Then
            load_Area()
        ElseIf rb_Date.Checked Then
            load_Date()
        End If

        th_LoadingWindow.Abort()

    End Sub

    Structure ItemList
        Dim partNo As String
        Dim partQty As Integer
        Dim lotQty As Integer
    End Structure

    Private Sub load_Date()

        Dim pmic_list() As ItemList = Nothing
        Dim rcd_list() As ItemList = Nothing

        grid_RejectList_Detail.Rows.Count = 2
        grid_RejectList_Detail.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_reject_material_ship(5" &
                                ", '" & Format(date_Reject_Start.Value, "yyyy-MM-dd 00:00:00") & "'" &
                                ", '" & Format(date_Reject_End.Value, "yyyy-MM-dd 23:59:59") & "'" &
                                ",null" &
                                ",null" &
                                ");"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = "N" & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          sqlDR("product") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("chip_qty") & vbTab &
                                          sqlDR("pmic_part_no") & vbTab &
                                          sqlDR("pmic_qty") & vbTab &
                                          sqlDR("rcd_part_no") & vbTab &
                                          sqlDR("rcd_qty")
            grid_RejectList_Detail.AddItem(insert_String)
            grid_RejectList_Detail.Rows(grid_RejectList_Detail.Rows.Count - 1).StyleNew.ForeColor = Color.Blue

            If IsNothing(pmic_list) Then
                ReDim pmic_list(0)
                pmic_list(0).partNo = sqlDR("pmic_part_no")
                pmic_list(0).partQty = sqlDR("pmic_qty")
                pmic_list(0).lotQty = 1
            Else
                Dim partFind As Boolean = False
                Dim findIndex As Integer = 0
                For i = 0 To pmic_list.Length - 1
                    If pmic_list(i).partNo.Equals(sqlDR("pmic_part_no")) Then
                        partFind = True
                        findIndex = i
                        Exit For
                    End If
                Next
                If partFind = True Then
                    pmic_list(findIndex).partQty += sqlDR("pmic_qty")
                    pmic_list(findIndex).lotQty += 1
                Else
                    ReDim Preserve pmic_list(pmic_list.Length)
                    pmic_list(pmic_list.Length - 1).partNo = sqlDR("pmic_part_no")
                    pmic_list(pmic_list.Length - 1).partQty = sqlDR("pmic_qty")
                    pmic_list(pmic_list.Length - 1).lotQty = 1
                End If
            End If

            If IsNothing(rcd_list) Then
                ReDim rcd_list(0)
                rcd_list(0).partNo = sqlDR("rcd_part_no")
                rcd_list(0).partQty = sqlDR("rcd_qty")
                rcd_list(0).lotQty = 1
            Else
                Dim partFind As Boolean = False
                Dim findIndex As Integer = 0
                For i = 0 To rcd_list.Length - 1
                    If rcd_list(i).partNo.Equals(sqlDR("rcd_part_no")) Then
                        partFind = True
                        findIndex = i
                        Exit For
                    End If
                Next
                If partFind = True Then
                    rcd_list(findIndex).partQty += sqlDR("rcd_qty")
                    rcd_list(findIndex).lotQty += 1
                Else
                    ReDim Preserve rcd_list(rcd_list.Length)
                    rcd_list(rcd_list.Length - 1).partNo = sqlDR("rcd_part_no")
                    rcd_list(rcd_list.Length - 1).partQty = sqlDR("rcd_qty")
                    rcd_list(rcd_list.Length - 1).lotQty = 1
                End If
            End If
        Loop
        sqlDR.Close()

        DBClose()

        grid_PartList.Redraw = False
        grid_PartList.Rows.Count = 1

        If Not IsNothing(pmic_list) Then
            For i = 0 To pmic_list.Length - 1
                If Not pmic_list(i).partNo.Equals("None") Then
                    grid_PartList.AddItem(grid_PartList.Rows.Count & vbTab &
                            "PMIC" & vbTab &
                            pmic_list(i).partNo & vbTab &
                            Format(pmic_list(i).lotQty, "#,##0") & vbTab &
                            Format(pmic_list(i).partQty, "#,##0"))
                End If
            Next
        End If

        If Not IsNothing(rcd_list) Then
            For i = 0 To rcd_list.Length - 1
                If Not rcd_list(i).partNo.Equals("None") Then
                    grid_PartList.AddItem(grid_PartList.Rows.Count & vbTab &
                            "RCD" & vbTab &
                            rcd_list(i).partNo & vbTab &
                            Format(rcd_list(i).lotQty, "#,##0") & vbTab &
                            Format(rcd_list(i).partQty, "#,##0"))
                End If
            Next
        End If

        grid_PartList.AutoSizeCols()
        grid_PartList.Redraw = True

        grid_RejectList_Detail.AutoSizeCols()
        grid_RejectList_Detail.Redraw = True

    End Sub

    Private Sub load_Area()

        Dim pmic_list() As ItemList = Nothing
        Dim rcd_list() As ItemList = Nothing

        If Trim(tb_Min_No.Text) = String.Empty And
                Trim(tb_Max_No.Text) = String.Empty Then
            MsgBox("범위를 입력하여 주십시오.",
                   MsgBoxStyle.Information,
                   form_Msgbox_String)
            Exit Sub
        End If

        If CInt(tb_Min_No.Text) > CInt(tb_Max_No.Text) Then
            MsgBox("범위가 잘 못 입력되었습니다.",
                   MsgBoxStyle.Information,
                   form_Msgbox_String)
            Exit Sub
        End If

        '범위 중 빠지는것을 확인하기 위해...
        Dim total_lot_qty As Integer = 0
        Dim load_qty As Integer = 0
        For i = CInt(tb_Min_No.Text) To CInt(tb_Max_No.Text)
            total_lot_qty += 1
        Next

        grid_RejectList_Detail.Rows.Count = 2
        grid_RejectList_Detail.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_reject_material_ship(1" &
                                ",null" &
                                ",null" &
                                ",'" & Format(CInt(tb_Min_No.Text), "00000") & "'" &
                                ",'" & Format(CInt(tb_Max_No.Text), "00000") & "'" &
                                ");"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = "N" & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          sqlDR("product") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("chip_qty") & vbTab &
                                          sqlDR("pmic_part_no") & vbTab &
                                          sqlDR("pmic_qty") & vbTab &
                                          sqlDR("rcd_part_no") & vbTab &
                                          sqlDR("rcd_qty")
            load_qty += 1
            grid_RejectList_Detail.AddItem(insert_String)
            grid_RejectList_Detail.Rows(grid_RejectList_Detail.Rows.Count - 1).StyleNew.ForeColor = Color.Blue

            If IsNothing(pmic_list) Then
                ReDim pmic_list(0)
                pmic_list(0).partNo = sqlDR("pmic_part_no")
                pmic_list(0).partQty = sqlDR("pmic_qty")
                pmic_list(0).lotQty = 1
            Else
                Dim partFind As Boolean = False
                Dim findIndex As Integer = 0
                For i = 0 To pmic_list.Length - 1
                    If pmic_list(i).partNo.Equals(sqlDR("pmic_part_no")) Then
                        partFind = True
                        findIndex = i
                        Exit For
                    End If
                Next
                If partFind = True Then
                    pmic_list(findIndex).partQty += sqlDR("pmic_qty")
                    pmic_list(findIndex).lotQty += 1
                Else
                    ReDim Preserve pmic_list(pmic_list.Length)
                    pmic_list(pmic_list.Length - 1).partNo = sqlDR("pmic_part_no")
                    pmic_list(pmic_list.Length - 1).partQty = sqlDR("pmic_qty")
                    pmic_list(pmic_list.Length - 1).lotQty = 1
                End If
            End If

            If IsNothing(rcd_list) Then
                ReDim rcd_list(0)
                rcd_list(0).partNo = sqlDR("rcd_part_no")
                rcd_list(0).partQty = sqlDR("rcd_qty")
                rcd_list(0).lotQty = 1
            Else
                Dim partFind As Boolean = False
                Dim findIndex As Integer = 0
                For i = 0 To rcd_list.Length - 1
                    If rcd_list(i).partNo.Equals(sqlDR("rcd_part_no")) Then
                        partFind = True
                        findIndex = i
                        Exit For
                    End If
                Next
                If partFind = True Then
                    rcd_list(findIndex).partQty += sqlDR("rcd_qty")
                    rcd_list(findIndex).lotQty += 1
                Else
                    ReDim Preserve rcd_list(rcd_list.Length)
                    rcd_list(rcd_list.Length - 1).partNo = sqlDR("rcd_part_no")
                    rcd_list(rcd_list.Length - 1).partQty = sqlDR("rcd_qty")
                    rcd_list(rcd_list.Length - 1).lotQty = 1
                End If
            End If
        Loop
        sqlDR.Close()

        DBClose()

        grid_PartList.Redraw = False
        grid_PartList.Rows.Count = 1

        For i = 0 To pmic_list.Length - 1
            If Not pmic_list(i).partNo.Equals("None") Then
                grid_PartList.AddItem(grid_PartList.Rows.Count & vbTab &
                        "PMIC" & vbTab &
                        pmic_list(i).partNo & vbTab &
                        Format(pmic_list(i).lotQty, "#,##0") & vbTab &
                        Format(pmic_list(i).partQty, "#,##0"))
            End If
        Next

        For i = 0 To rcd_list.Length - 1
            If Not rcd_list(i).partNo.Equals("None") Then
                grid_PartList.AddItem(grid_PartList.Rows.Count & vbTab &
                        "RCD" & vbTab &
                        rcd_list(i).partNo & vbTab &
                        Format(rcd_list(i).lotQty, "#,##0") & vbTab &
                        rcd_list(i).partQty)
            End If
        Next

        grid_PartList.AutoSizeCols()
        grid_PartList.Redraw = True

        grid_RejectList_Detail.AutoSizeCols()
        grid_RejectList_Detail.Redraw = True

        If Not total_lot_qty = load_qty Then
            MsgBox("지정된 범위 숫자 갯수와 불러온 데이터 갯수가 다릅니다." & vbCrLf & "누락된 데이터를 확인하여 주십시오.",
                   MsgBoxStyle.Information,
                   form_Msgbox_String)
        End If

        Dim nowYJNo As String = "Y" & Format(CInt(tb_Min_No.Text), "00000")
        Dim findYJNo As New List(Of String)()
        Dim noExistYJNo As New List(Of String)()

        For i = CInt(tb_Min_No.Text) To CInt(tb_Max_No.Text)
            nowYJNo = "Y" & Format(i, "00000")
            findYJNo.Add(nowYJNo)
        Next

        For i = 0 To findYJNo.Count - 1
            If grid_RejectList_Detail.FindRow(findYJNo(i).ToString, 2, 1, True) = -1 Then
                noExistYJNo.Add(findYJNo(i).ToString)
            End If
        Next

        Dim showMessage As String = String.Empty
        For i = 0 To noExistYJNo.Count - 1
            If showMessage = String.Empty Then
                showMessage = noExistYJNo(i).ToString
            Else
                showMessage += ", " & noExistYJNo(i).ToString
            End If
        Next

        TextBox2.Text = showMessage

    End Sub

    Private Sub btn_Save_Click(sender As Object, e As EventArgs) Handles btn_Save.Click

        If grid_RejectList_Detail.Rows.Count = 2 Then
            MsgBox("저장할 내용이 없습니다.",
                   MsgBoxStyle.Information,
                   form_Msgbox_String)
            Exit Sub
        End If

        For i = 1 To grid_PartList.Rows.Count - 1
            If Trim(grid_PartList(i, 5)) = String.Empty Then
                MsgBox("행 : " & i & vbCrLf &
                    "Part No. : " & grid_PartList(i, 2) & "의" & vbCrLf &
                    "무게를 입력하여 주십시오.",
                       MsgBoxStyle.Information,
                       form_Msgbox_String)
                Exit Sub
            End If
        Next

        If MsgBox("저장 하시겠습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                  form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Saving...")

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try
            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            For i = 2 To grid_RejectList_Detail.Rows.Count - 1
                Select Case grid_RejectList_Detail(i, 0).ToString
                    Case "N"
                        strSQL += "insert into reject_material_ship_information(ship_no, lot_no, yj_no"
                        strSQL += ", ship_date, writer, write_date, first_write_date, ship_note) values("
                        strSQL += "'" & TextBox1.Text & "'"
                        strSQL += ",'" & grid_RejectList_Detail(i, 3) & "'"
                        strSQL += ",'" & grid_RejectList_Detail(i, 1) & "'"
                        strSQL += ",'" & Format(DateTimePicker3.Value, "yyyy-MM-dd") & "'"
                        strSQL += ",'" & loginID & "'"
                        strSQL += ",'" & write_date & "'"
                        strSQL += ",'" & write_date & "'"
                        strSQL += ",'" & grid_RejectList_Detail(i, 9) & "');"
                    Case "D"
                        strSQL += "delete from reject_material_ship_information"
                        strSQL += " where ship_no = '" & TextBox1.Text & "'"
                        strSQL += " and lot_no = '" & grid_RejectList_Detail(i, 3) & "';"
                End Select
            Next

            strSQL += "delete from reject_material_ship_summary where ship_no = '" & TextBox1.Text & "';"
            For i = 1 To grid_PartList.Rows.Count - 1
                strSQL += "insert into reject_material_ship_summary(ship_no, part_division,"
                strSQL += "part_no, lot_qty, part_qty, weight, ship_note) values("
                strSQL += "'" & TextBox1.Text & "'"
                strSQL += ",'" & grid_PartList(i, 1) & "'"
                strSQL += ",'" & grid_PartList(i, 2) & "'"
                strSQL += ",'" & CInt(grid_PartList(i, 3)) & "'"
                strSQL += ",'" & CInt(grid_PartList(i, 4)) & "'"
                strSQL += ",'" & CDbl(grid_PartList(i, 5)) & "'"
                strSQL += ",'" & grid_PartList(i, 6) & "'"
                strSQL += ");"
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

        th_LoadingWindow.Abort()
        Thread.Sleep(100)
        MsgBox("폐기자재 반납 내용을 저장 하였습니다.", MsgBoxStyle.Information, form_Msgbox_String)

        grid_RejectList_Detail.Rows.Count = 2
        grid_PartList.Rows.Count = 1
        btn_Save.Enabled = False
        TextBox1.Text = String.Empty
        TextBox2.Text = "여기는 누락된 데이터가 표시되는 곳입니다."
        date_Reject_Start.Enabled = False
        date_Reject_End.Enabled = False
        tb_Min_No.Enabled = False
        tb_Max_No.Enabled = False
        btn_PMIC_Print.Enabled = False

    End Sub

    Private Sub grid_RejectList_Detail_MouseClick(sender As Object, e As MouseEventArgs) Handles grid_RejectList_Detail.MouseClick

        If e.Button = MouseButtons.Right And grid_RejectList_Detail.MouseRow > 0 Then
            grid_RejectList_Detail.Row = grid_RejectList_Detail.MouseRow
            cms_Menu.Show(grid_RejectList_Detail, New Point(e.X, e.Y))
        End If

    End Sub

    Private Sub btn_RowDelete_Click(sender As Object, e As EventArgs) Handles btn_RowDelete.Click

        If grid_RejectList_Detail(grid_RejectList_Detail.Row, 0).ToString = "N" Then
            grid_RejectList_Detail.RemoveItem(grid_RejectList_Detail.Row)
        Else
            grid_RejectList_Detail(grid_RejectList_Detail.Row, 0) = "D"
        End If

        grid_RejectList_Detail.Redraw = False

        For i = 2 To grid_RejectList_Detail.Rows.Count - 1
            If IsNumeric(grid_RejectList_Detail(i, 0)) Then
                grid_RejectList_Detail(i, 0) = i - 1
                grid_RejectList_Detail.Rows(i).StyleNew.ForeColor = Color.Black
            Else
                If grid_RejectList_Detail(i, 0).ToString = "N" Then
                    grid_RejectList_Detail.Rows(i).StyleNew.ForeColor = Color.Blue
                ElseIf grid_RejectList_Detail(i, 0).ToString = "D" Then
                    grid_RejectList_Detail.Rows(i).StyleNew.ForeColor = Color.Gray
                End If
            End If
        Next

        grid_RejectList_Detail.AutoSizeCols()

        grid_RejectList_Detail.Redraw = True

        Dim pmic_list() As ItemList = Nothing
        Dim rcd_list() As ItemList = Nothing

        For i = 2 To grid_RejectList_Detail.Rows.Count - 1
            If IsNothing(pmic_list) Then
                ReDim pmic_list(0)
                pmic_list(0).partNo = grid_RejectList_Detail(i, 5)
                pmic_list(0).partQty = grid_RejectList_Detail(i, 6)
                pmic_list(0).lotQty = 1
            Else
                Dim partFind As Boolean = False
                Dim findIndex As Integer = 0
                For j = 0 To pmic_list.Length - 1
                    If pmic_list(j).partNo.Equals(grid_RejectList_Detail(i, 5)) Then
                        partFind = True
                        findIndex = j
                        Exit For
                    End If
                Next
                If partFind = True Then
                    pmic_list(findIndex).partQty += grid_RejectList_Detail(i, 6)
                    pmic_list(findIndex).lotQty += 1
                Else
                    ReDim Preserve pmic_list(pmic_list.Length)
                    pmic_list(pmic_list.Length - 1).partNo = grid_RejectList_Detail(i, 5)
                    pmic_list(pmic_list.Length - 1).partQty = grid_RejectList_Detail(i, 6)
                    pmic_list(pmic_list.Length - 1).lotQty = 1
                End If
            End If

            If IsNothing(rcd_list) Then
                ReDim rcd_list(0)
                rcd_list(0).partNo = grid_RejectList_Detail(i, 7)
                rcd_list(0).partQty = grid_RejectList_Detail(i, 8)
                rcd_list(0).lotQty = 1
            Else
                Dim partFind As Boolean = False
                Dim findIndex As Integer = 0
                For j = 0 To rcd_list.Length - 1
                    If rcd_list(j).partNo.Equals(grid_RejectList_Detail(i, 7)) Then
                        partFind = True
                        findIndex = j
                        Exit For
                    End If
                Next
                If partFind = True Then
                    rcd_list(findIndex).partQty += grid_RejectList_Detail(i, 8)
                    rcd_list(findIndex).lotQty += 1
                Else
                    ReDim Preserve rcd_list(rcd_list.Length)
                    rcd_list(rcd_list.Length - 1).partNo = grid_RejectList_Detail(i, 7)
                    rcd_list(rcd_list.Length - 1).partQty = grid_RejectList_Detail(i, 8)
                    rcd_list(rcd_list.Length - 1).lotQty = 1
                End If
            End If
        Next

        grid_PartList.Redraw = False
        grid_PartList.Rows.Count = 1

        For i = 0 To pmic_list.Length - 1
            If Not pmic_list(i).partNo.Equals("None") Then
                grid_PartList.AddItem(grid_PartList.Rows.Count & vbTab &
                        "PMIC" & vbTab &
                        pmic_list(i).partNo & vbTab &
                        Format(pmic_list(i).lotQty, "#,##0") & vbTab &
                        Format(pmic_list(i).partQty, "#,##0"))
            End If
        Next

        For i = 0 To rcd_list.Length - 1
            If Not rcd_list(i).partNo.Equals("None") Then
                grid_PartList.AddItem(grid_PartList.Rows.Count & vbTab &
                        "RCD" & vbTab &
                        rcd_list(i).partNo & vbTab &
                        Format(rcd_list(i).lotQty, "#,##0") & vbTab &
                        rcd_list(i).partQty)
            End If
        Next

        grid_PartList.AutoSizeCols()
        grid_PartList.Redraw = True

    End Sub

    Private Sub btn_Search_Click(sender As Object, e As EventArgs) Handles btn_Search.Click

        thread_LoadingFormStart()

        TextBox1.Text = String.Empty
        date_Reject_Start.Enabled = False
        date_Reject_End.Enabled = False
        tb_Max_No.Enabled = False
        tb_Min_No.Enabled = False
        btn_PMIC_Print.Enabled = False

        grid_PartList.Rows.Count = 1
        grid_RejectList_Detail.Rows.Count = 2
        grid_RejectList.Rows.Count = 1
        grid_RejectList.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_reject_material_ship(2" &
                                ",'" & Format(DateTimePicker1.Value, "yyyy-MM-dd 00:00:00") & "'" &
                                ",'" & Format(DateTimePicker2.Value, "yyyy-MM-dd 23:59:59") & "'" &
                                ",'" & tb_S_Lot_No.Text & "'" &
                                ",'" & tb_S_YJ_No.Text & "')"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_RejectList.Rows.Count & vbTab &
                                          Format(sqlDR("ship_date"), "yyyy-MM-dd") & vbTab &
                                          sqlDR("ship_no") & vbTab &
                                          Format(sqlDR("chip_qty"), "#,##0")
            grid_RejectList.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        grid_RejectList.AutoSizeCols()
        grid_RejectList.Redraw = True

        th_LoadingWindow.Abort()

    End Sub

    Private Sub grid_RejectList_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles grid_RejectList.MouseDoubleClick

        Dim sel_row As Integer = grid_RejectList.MouseRow

        If e.Button = MouseButtons.Left And sel_row > 0 Then

            thread_LoadingFormStart()

            grid_RejectList_Detail.Rows.Count = 2
            grid_RejectList_Detail.Redraw = False

            TextBox1.Text = grid_RejectList(sel_row, 2)
            btn_Save.Enabled = True

            DBConnect()

            Dim strSQL As String = "call sp_reject_material_ship(3" &
                                    ",null" &
                                    ",null" &
                                    ",'" & grid_RejectList(sel_row, 2) & "'" &
                                    ",null)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                Dim insert_String As String = grid_RejectList_Detail.Rows.Count - 1 & vbTab &
                                              sqlDR("yj_no") & vbTab &
                                              sqlDR("product") & vbTab &
                                              sqlDR("lot_no") & vbTab &
                                              sqlDR("chip_qty") & vbTab &
                                              sqlDR("pmic_part_no") & vbTab &
                                              sqlDR("pmic_qty") & vbTab &
                                              sqlDR("rcd_part_no") & vbTab &
                                              sqlDR("rcd_qty") & vbTab &
                                              sqlDR("ship_note")
                grid_RejectList_Detail.AddItem(insert_String)
            Loop
            sqlDR.Close()

            grid_RejectList_Detail.AutoSizeCols()
            grid_RejectList_Detail.Redraw = True

            grid_PartList.Redraw = False
            grid_PartList.Rows.Count = 1

            strSQL = "call sp_reject_material_ship(6" &
                      ",null" &
                      ",null" &
                      ",'" & grid_RejectList(sel_row, 2) & "'" &
                      ",'')"

            sqlCmd = New MySqlCommand(strSQL, dbConnection1)
            sqlDR = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                Dim insert_String As String = grid_PartList.Rows.Count & vbTab &
                                              sqlDR("part_division") & vbTab &
                                              sqlDR("part_no") & vbTab &
                                              Format(sqlDR("lot_qty"), "#,##0") & vbTab &
                                              Format(sqlDR("part_qty"), "#,##0") & vbTab &
                                              Format(sqlDR("weight"), "#,##0.00") & vbTab &
                                              sqlDR("ship_note")
                grid_PartList.AddItem(insert_String)
            Loop
            sqlDR.Close()

            DBClose()

            grid_PartList.AutoSizeCols()
            grid_PartList.Redraw = True

            btn_PMIC_Print.Enabled = True

            th_LoadingWindow.Abort()
        End If

    End Sub

    Private Sub Load_ShipInfo(ByVal slip_no As String)

        Dim strSQL As String = String.Empty

        'MySQL DB에서 정보를 불러오기전에 기존 내용 삭제
        Mdbconnect()

        Dim sqlTran_MDB As OleDb.OleDbTransaction
        Dim sqlCmd_MDB As OleDb.OleDbCommand

        sqlTran_MDB = mdbConnection1.BeginTransaction

        Try
            '기존 저장 데이터를 삭제
            strSQL = "delete from material_ship_information;"

            sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
            sqlCmd_MDB.Transaction = sqlTran_MDB
            sqlCmd_MDB.ExecuteNonQuery()

            strSQL = "delete from material_ship_information_detail;"

            sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
            sqlCmd_MDB.Transaction = sqlTran_MDB
            sqlCmd_MDB.ExecuteNonQuery()

            strSQL = "delete from material_ship_information_each_part;"

            sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
            sqlCmd_MDB.Transaction = sqlTran_MDB
            sqlCmd_MDB.ExecuteNonQuery()

            sqlTran_MDB.Commit()

        Catch ex As OleDb.OleDbException
            sqlTran_MDB.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        '새로운 데이터를 MySQL에서 가져온다.
        sqlTran_MDB = mdbConnection1.BeginTransaction

        Try
            DBConnect()

            strSQL = "call sp_reject_material_ship(4"
            strSQL += ", null"
            strSQL += ", null"
            strSQL += ",'" & slip_no & "'"
            strSQL += ",null)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                strSQL = "insert into material_ship_information(ship_no, ship_date, write_date) values"
                strSQL += "('" & sqlDR("ship_no") & "'"
                strSQL += ",'" & sqlDR("ship_date") & "'"
                strSQL += ",'" & sqlDR("first_write_date") & "');"

                sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
                sqlCmd_MDB.Transaction = sqlTran_MDB
                sqlCmd_MDB.ExecuteNonQuery()
            Loop
            sqlDR.Close()

            strSQL = "call sp_reject_material_ship(3"
            strSQL += ", null"
            strSQL += ", null"
            strSQL += ",'" & slip_no & "'"
            strSQL += ",null)"

            sqlCmd = New MySqlCommand(strSQL, dbConnection1)
            sqlDR = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                strSQL = "insert into material_ship_information_detail(ship_no, yj_no, product, lot_no"
                strSQL += ", chip_qty, pmic_part_no, pmic_qty, rcd_part_no, rcd_qty, ship_note) values"
                strSQL += "('" & sqlDR("ship_no") & "'"
                strSQL += ",'" & sqlDR("yj_no") & "'"
                strSQL += ",'" & sqlDR("product") & "'"
                strSQL += ",'" & sqlDR("lot_no") & "'"
                strSQL += ",'" & sqlDR("chip_qty") & "'"
                strSQL += ",'" & sqlDR("pmic_part_no") & "'"
                strSQL += ",'" & sqlDR("pmic_qty") & "'"
                strSQL += ",'" & sqlDR("rcd_part_no") & "'"
                strSQL += ",'" & sqlDR("rcd_qty") & "'"
                strSQL += ",'" & sqlDR("ship_note") & "');"

                sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
                sqlCmd_MDB.Transaction = sqlTran_MDB
                sqlCmd_MDB.ExecuteNonQuery()
            Loop
            sqlDR.Close()

            strSQL = "call sp_reject_material_ship(6"
            strSQL += ", null"
            strSQL += ", null"
            strSQL += ",'" & slip_no & "'"
            strSQL += ",null)"

            sqlCmd = New MySqlCommand(strSQL, dbConnection1)
            sqlDR = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                strSQL = "insert into material_ship_information_each_part(ship_no, part_division, part_no, lot_qty, part_qty, weight, ship_note) values"
                strSQL += "('" & sqlDR("ship_no") & "'"
                strSQL += ",'" & sqlDR("part_division") & "'"
                strSQL += ",'" & sqlDR("part_no") & "'"
                strSQL += ",'" & sqlDR("lot_qty") & "'"
                strSQL += ",'" & sqlDR("part_qty") & "'"
                strSQL += ",'" & sqlDR("weight") & "'"
                strSQL += ",'" & sqlDR("ship_note") & "');"

                sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
                sqlCmd_MDB.Transaction = sqlTran_MDB
                sqlCmd_MDB.ExecuteNonQuery()
            Loop
            sqlDR.Close()

            DBClose()

            sqlTran_MDB.Commit()

        Catch ex As OleDb.OleDbException
            sqlTran_MDB.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        MDBClose()

    End Sub

    Private Sub ReportLoad()

        '크리스탈 레포트에서 인쇄하는 규칙
        Try
            Dim rptPath As String = Application.StartupPath & "\Reports\rpt_Reject_Material_Ship_Summary.rpt"

            Dim ds As DataSet = New DataSet

            Dim connection As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath + "\TempDB\TempDB.mdb" & ";Jet OLEDB:Database Password='dbwlspark'")
            connection.Open()

            Dim strQuery As String = " SELECT `material_ship_information`.`ship_no`, `material_ship_information`.`write_date`,"
            strQuery += " `material_ship_information_each_part`.`part_no`, `material_ship_information_each_part`.`part_qty`,"
            strQuery += " `material_ship_information_each_part`.`lot_qty`"
            strQuery += " FROM   `material_ship_information` `material_ship_information` INNER JOIN `material_ship_information_each_part` `material_ship_information_each_part`"
            strQuery += " ON `material_ship_information`.`ship_no`=`material_ship_information_each_part`.`ship_no`"
            strQuery += " ORDER BY `material_ship_information_each_part`.`part_no`"

            'Console.WriteLine(strQuery)

            Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(strQuery, connection)
            da.Fill(ds)

            Dim rDOC As ReportDocument = New ReportDocument

            rDOC.Load(rptPath)
            rDOC.SetDataSource(ds)
            'rDOC.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape '용지방향 설정
            frm_Report_Print.CrystalReportViewer1.ReportSource = Nothing
            frm_Report_Print.CrystalReportViewer1.ReportSource = rDOC

            If MsgBox("인쇄 미리보기를 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.Yes Then
                frm_Report_Print.ShowDialog()
            Else
                rDOC.PrintToPrinter(3, True, 0, 0)
            End If

            connection.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
        End Try

    End Sub

    Private Sub tb_Min_No_TextChanged(sender As Object, e As EventArgs) Handles tb_Min_No.TextChanged

    End Sub

    Private Sub date_Reject_Start_ValueChanged(sender As Object, e As EventArgs) Handles date_Reject_Start.ValueChanged

        rb_Date.Checked = True

    End Sub

    Private Sub date_Reject_End_ValueChanged(sender As Object, e As EventArgs) Handles date_Reject_End.ValueChanged

        rb_Date.Checked = True

    End Sub

    Private Sub btn_PMIC_Print_Click(sender As Object, e As EventArgs) Handles btn_PMIC_Print.Click

        Load_ShipInfo(TextBox1.Text)
        ReportLoad()

    End Sub

    Private Sub grid_PartList_RowColChange(sender As Object, e As EventArgs) Handles grid_PartList.RowColChange

        Select Case grid_PartList.Col
            Case 5, 6
                grid_PartList.AllowEditing = True
            Case Else
                grid_PartList.AllowEditing = False
        End Select

    End Sub

    Private Sub grid_PartList_AfterEdit(sender As Object, e As RowColEventArgs) Handles grid_PartList.AfterEdit

        Select Case e.Col
            Case 5
                If IsNumeric(grid_PartList(e.Row, e.Col)) Then
                    grid_PartList(e.Row, e.Col) = Format(CDbl(grid_PartList(e.Row, e.Col)), "#,##0.00")
                Else
                    MsgBox("무게는 숫자로만 입력하여 주십시오.",
                           MsgBoxStyle.Information,
                           form_Msgbox_String)
                    grid_PartList(e.Row, e.Col) = String.Empty
                End If
        End Select

    End Sub

    Private Sub grid_PartList_MouseMove(sender As Object, e As MouseEventArgs) Handles grid_PartList.MouseMove, grid_RejectList.MouseMove, grid_RejectList_Detail.MouseMove

        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then

            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub

    Private Sub grid_LotList_LostFocus(sender As Object, e As EventArgs) Handles grid_PartList.LostFocus,
                                                                             grid_RejectList.LostFocus,
                                                                             grid_RejectList_Detail.LostFocus

        frm_Main.lb_Status.Text = String.Empty

    End Sub

    Private Sub grid_RejectList_Detail_KeyDown(sender As Object, e As KeyEventArgs) Handles grid_RejectList_Detail.KeyDown,
                                                                                            grid_RejectList.KeyDown,
                                                                                            grid_PartList.KeyDown

        If e.Shift Then
            md_ETC.cellCal(sender, Nothing)
        Else
            frm_Main.lb_Status.Text = String.Empty
        End If

    End Sub
End Class
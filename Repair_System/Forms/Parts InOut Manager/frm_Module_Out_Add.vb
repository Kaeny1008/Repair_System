Imports System.Threading
Imports C1.Win.C1FlexGrid
Imports CrystalDecisions.CrystalReports.Engine
Imports MySql.Data.MySqlClient

Public Class frm_Module_Out_Add

    Dim form_msgbox_string As String = "출고 등록(전표생성)"
    Dim grid_status As String = String.Empty

    Private Sub frm_Module_Out_Add_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Setting_Grid()

        DateTimePicker1.Value = Format(Now, "yyyy-MM-01 00:00:00")
        DateTimePicker2.Value = Format(Now, "yyyy-MM-dd 23:59:59")

    End Sub

    Private Sub Setting_Grid()

        With grid_ShipList
            .AllowEditing = False
            .AllowFiltering = False
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 9
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
            Mrng = .GetCellRange(0, 5, 1, 7)
            Mrng.Data = "수리정보"
            Mrng = .GetCellRange(0, 8, 1, 8)
            Mrng.Data = "비고"
            grid_ShipList(1, 2) = "Product Code"
            grid_ShipList(1, 3) = "Lot No."
            grid_ShipList(1, 4) = "수량(EA)"
            grid_ShipList(1, 5) = "구분"
            grid_ShipList(1, 6) = "Part No."
            grid_ShipList(1, 7) = "사용수량"
            '.Styles("Normal").WordWrap = True
            '.Cols(11).StyleNew.WordWrap = True

            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 1).TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ExtendLastCol = True
        End With

        With grid_OutList
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
            grid_OutList(0, 0) = "No"
            grid_OutList(0, 1) = "출고일자"
            grid_OutList(0, 2) = "Slip No."
            grid_OutList(0, 3) = "Lot 수"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

    End Sub

    Private Sub btn_New_Kitting_Click(sender As Object, e As EventArgs) Handles btn_New_Kitting.Click

        grid_status = "New"
        btn_Save.Enabled = True
        grid_ShipList.Redraw = False
        grid_ShipList.Rows.Count = 2

        tb_SlipNo.Text = SlipNo_Make(Now)

        grid_ShipList.AddItem("N")
        grid_ShipList.Rows(grid_ShipList.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
        grid_ShipList.Focus()
        grid_ShipList.Select(2, 1)
        grid_ShipList.Redraw = True
        Label10.Text = "총 출고 Lot수 : " & grid_ShipList.Rows.Count - 2 & " Lot"

        'btn_TempLoad.Enabled = True
        'btn_TempSave.Enabled = True
        btn_PDA_TempLoad.Enabled = True

    End Sub

    Private Sub grid_ShipList_RowColChange(sender As Object, e As EventArgs) Handles grid_ShipList.RowColChange

        Select Case grid_ShipList.Col
            Case 1
                If grid_status = "New" Then
                    grid_ShipList.AllowEditing = True
                Else
                    grid_ShipList.AllowEditing = False
                End If
            Case 8
                grid_ShipList.AllowEditing = True
            Case 3
                If CheckBox1.Checked Then
                    grid_ShipList.AllowEditing = True
                Else
                    grid_ShipList.AllowEditing = False
                End If
            Case Else
                grid_ShipList.AllowEditing = False
        End Select

    End Sub

    Private Sub grid_ShipList_KeyDown(sender As Object, e As KeyEventArgs) Handles grid_ShipList.KeyDown

        Select Case e.KeyCode
            Case 112 'F1
                grid_ShipList.AddItem("N")
                grid_ShipList.Select(grid_ShipList.Rows.Count - 1, 1)
                grid_ShipList.Rows(grid_ShipList.Row).StyleNew.ForeColor = Color.Blue
                Label10.Text = "총 출고 Lot수 : " & grid_ShipList.Rows.Count - 2 & " Lot"
            Case 114 'F3
                If grid_ShipList(grid_ShipList.Row, 0) = "N" Then
                    grid_ShipList.RemoveItem(grid_ShipList.Row)
                Else
                    grid_ShipList(grid_ShipList.Row, 0) = "D"
                    grid_ShipList.Rows(grid_ShipList.Row).StyleNew.ForeColor = Color.Gray
                End If
                Label10.Text = "총 출고 Lot수 : " & grid_ShipList.Rows.Count - 2 & " Lot"
            Case Else

        End Select

        If e.Shift Then
            md_ETC.cellCal(sender, Nothing)
        Else
            frm_Main.lb_Status.Text = String.Empty
        End If

    End Sub

    Private Sub grid_ShipList_AfterEdit(sender As Object, e As RowColEventArgs) Handles grid_ShipList.AfterEdit

        If Trim(grid_ShipList(e.Row, e.Col)).Equals(String.Empty) Then Exit Sub
        If Strings.Right(grid_ShipList(e.Row, e.Col), 1) = "\" Then
            MsgBox("마지막 문자는 '\'가 될수 없습니다.",
                   MsgBoxStyle.Information,
                   form_msgbox_string)
            Exit Sub
        End If

        grid_ShipList.Redraw = False

        If Not grid_ShipList(e.Row, 0) = "N" Then
            grid_ShipList(e.Row, 0) = "M"
            grid_ShipList.Rows(e.Row).StyleNew.ForeColor = Color.Red
        End If

        Select Case e.Col
            Case 1
                grid_ShipList(e.Row, 1) = grid_ShipList(e.Row, 1).ToString.ToUpper
                'Lot의 상태( WorkSite Moving )를 확인한다  
                Dim lotStatus As String = Load_YJNo_Status(grid_ShipList(e.Row, 1))
                If Not lotStatus = "Moving to WorkSite" And
                    Not lotStatus = "PMIC Working Completed" And
                    Not lotStatus = "RCD Working Completed" Then
                    MsgBox("YJ No. : " & grid_ShipList(e.Row, 1) & " 는(은)" &
                           vbCrLf & "현장 출고상태가 아닙니다.",
                           MsgBoxStyle.Information,
                           form_msgbox_string)
                    grid_ShipList(e.Row, 1) = String.Empty
                    grid_ShipList.AutoSizeCols()
                    grid_ShipList.AutoSizeRows(2, 0, grid_ShipList.Rows.Count - 1, grid_ShipList.Cols.Count - 1, 0, AutoSizeFlags.None)
                    grid_ShipList.Redraw = True
                    Exit Sub
                End If
                Load_Lot_Status(grid_ShipList(e.Row, 1), e.Row)
                Load_Repair_Data(grid_ShipList(e.Row, 1), e.Row, False)
            Case 3
                grid_ShipList(e.Row, 3) = grid_ShipList(e.Row, 3).ToString.ToUpper
                If Not Load_Lot_Status(grid_ShipList(e.Row, 3)) = "Ready" Then
                    MsgBox("Lot No. : " & grid_ShipList(e.Row, 3) & " 는(은)" &
                           vbCrLf & "수입검사 대기 품목이 아닙니다.",
                           MsgBoxStyle.Information,
                           form_msgbox_string)
                    grid_ShipList(e.Row, 3) = String.Empty
                    grid_ShipList.AutoSizeCols()
                    grid_ShipList.AutoSizeRows(2, 0, grid_ShipList.Rows.Count - 1, grid_ShipList.Cols.Count - 1, 0, AutoSizeFlags.None)
                    grid_ShipList.Redraw = True
                    Exit Sub
                Else
                    Load_Lot_Status(grid_ShipList(e.Row, 3), e.Row)
                End If
        End Select

        grid_ShipList.AutoSizeCols()
        grid_ShipList.AutoSizeRows(2, 0, grid_ShipList.Rows.Count - 1, grid_ShipList.Cols.Count - 1, 0, AutoSizeFlags.None)
        grid_ShipList.Redraw = True

    End Sub

    Private Function Load_Lot_Status(ByVal lot_no As String) As String

        Load_Lot_Status = String.Empty

        DBConnect()

        Dim strSql As String = "call sp_module_ship(6" &
            ", null" &
            ", null" &
            ", '" & lot_no & "'" &
            ", null)"
        Dim sqlCmd As New MySqlCommand(strSql, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Load_Lot_Status = sqlDR("lot_status")
        Loop

        sqlDR.Close()
        DBClose()

        Return Load_Lot_Status

    End Function

    Private Function Load_YJNo_Status(ByVal yj_no As String) As String

        Load_YJNo_Status = String.Empty

        DBConnect()

        Dim strSql As String = "call sp_module_ship(0" &
            ", null" &
            ", null" &
            ", '" & yj_no & "'" &
            ", null)"
        Dim sqlCmd As New MySqlCommand(strSql, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Load_YJNo_Status = sqlDR("lot_status")
        Loop

        sqlDR.Close()
        DBClose()

        Return Load_YJNo_Status

    End Function

    Private Sub Load_Lot_Status(ByVal yj_no As String, ByVal editRow As Integer)

        DBConnect()

        Dim strSql As String = "call sp_module_ship(0" &
            ", null" &
            ", null" &
            ", '" & yj_no & "'" &
            ", null)"
        Dim sqlCmd As New MySqlCommand(strSql, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            grid_ShipList(editRow, 1) = sqlDR("yj_no")
            grid_ShipList(editRow, 2) = sqlDR("product")
            grid_ShipList(editRow, 3) = sqlDR("lot_no")
            grid_ShipList(editRow, 4) = sqlDR("chip_qty")
        Loop

        sqlDR.Close()
        DBClose()

    End Sub

    Private Sub Load_Repair_Data(ByVal yj_no As String, ByVal editRow As Integer, ByVal nowLoad As Boolean)

        DBConnect()

        Dim strSql As String = "call sp_module_ship(1" &
            ", null" &
            ", null" &
            ", '" & yj_no & "'" &
            ", null)"
        Dim sqlCmd As New MySqlCommand(strSql, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader
        Dim read_Count As Integer = 0

        Do While sqlDR.Read
            If sqlDR("pfq_doe") = "False" Then
                If read_Count = 0 Then
                    grid_ShipList(editRow, 5) = sqlDR("part_division")
                    grid_ShipList(editRow, 6) = sqlDR("part_no")
                    grid_ShipList(editRow, 7) = CStr(sqlDR("chip_qty"))
                    If nowLoad = False Then
                        If sqlDR("lot_option") Like "*O" Then
                            grid_ShipList(editRow, 8) = "PMIC Repair"
                        ElseIf sqlDR("lot_option") Like "*Q" Then
                            grid_ShipList(editRow, 8) = "PMIC+RCD Repair"
                        End If
                    End If
                Else
                    grid_ShipList(editRow, 5) += vbCrLf & sqlDR("part_division")
                    grid_ShipList(editRow, 6) += vbCrLf & sqlDR("part_no")
                    grid_ShipList(editRow, 7) += vbCrLf & CStr(sqlDR("chip_qty"))
                End If
                read_Count += 1
            End If
        Loop

        sqlDR.Close()
        DBClose()

    End Sub

    Private Function SlipNo_Make(ByVal findDate As Date) As String

        SlipNo_Make = String.Empty

        DBConnect()

        Dim strSql As String = "select slip_no from module_ship_information" &
                                " where first_write_date Like '" & Format(findDate, "yyyy-MM-dd") & "%'" &
                                " order by slip_no desc limit 1"
        Dim sqlCmd As New MySqlCommand(strSql, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            SlipNo_Make = sqlDR("slip_no")
        Loop

        sqlDR.Close()
        DBClose()

        If SlipNo_Make = String.Empty Then SlipNo_Make = "UJ" & Format(Now, "yyMMdd") & Format(0, "00")
        SlipNo_Make = SlipNo_Make.Substring(8)
        SlipNo_Make = "UJ" & Format(CDate(findDate), "yyMMdd") & Format(SlipNo_Make + 1, "00")

        Return SlipNo_Make

    End Function

    Private Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        thread_LoadingFormStart()

        btn_TempLoad.Enabled = False
        btn_TempSave.Enabled = False
        btn_PDA_TempLoad.Enabled = False
        btn_Save.Enabled = False
        tb_SlipNo.Text = String.Empty
        grid_ShipList.Rows.Count = 2

        grid_OutList.Redraw = False
        grid_OutList.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call sp_module_ship(5" &
            ", '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "'" &
            ", '" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker2.Value), "yyyy-MM-dd") & "'" &
            ", '" & TB_S_YJ_No.Text & "'" &
            ", null)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_OutList.Rows.Count & vbTab &
                                          Format(sqlDR("ship_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          sqlDR("slip_no") & vbTab &
                                          Format(sqlDR("ship_lot_count"), "#,##0")
            grid_OutList.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        grid_OutList.AutoSizeCols()
        grid_OutList.Redraw = True

        thread_LoadingFormEnd

    End Sub

    Private Sub btn_Save_Click(sender As Object, e As EventArgs) Handles btn_Save.Click

        If MsgBox("저장 하시겠습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                  form_msgbox_string) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Saving...")

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = "delete from module_ship_information_pda_temp;"

        sqlTran = dbConnection1.BeginTransaction

        Try
            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            '라벨 발행 기록
            For i = 1 To grid_ShipList.Rows.Count - 1
                If grid_ShipList(i, 0) = "N" Then
                    strSQL += "insert into module_ship_information(slip_no, lot_no, ship_date, writer, write_date"
                    strSQL += ", first_write_date, ship_note) values("
                    strSQL += "'" & tb_SlipNo.Text & "'"
                    strSQL += ",'" & grid_ShipList(i, 3) & "'"
                    strSQL += ",'" & write_date & "'"
                    strSQL += ",'" & loginID & "'"
                    strSQL += ",'" & write_date & "'"
                    strSQL += ",'" & write_date & "'"
                    strSQL += ",'" & grid_ShipList(i, 8) & "');"

                    strSQL += "update basic_lot_information set lot_status = 'Completed'"
                    strSQL += " where lot_no = '" & grid_ShipList(i, 3) & "';"
                ElseIf grid_ShipList(i, 0) = "M" Then
                    strSQL += "update module_ship_information set ship_note = "
                    strSQL += "'" & grid_ShipList(i, 8) & "'"
                    strSQL += " where lot_no = '" & grid_ShipList(i, 3) & "';"
                ElseIf grid_ShipList(i, 0) = "D" Then
                    strSQL += "update basic_lot_information set lot_status = 'WorkSite Moving'" '현재는 WorkSite Moving이 출고 이전단계이므로 차후 조정하여야 함.
                    strSQL += " where lot_no = '" & grid_ShipList(i, 3) & "';"
                    strSQL += "delete from module_ship_information"
                    strSQL += " where lot_no = '" & grid_ShipList(i, 3) & "'"
                    strSQL += " and slip_no = '" & tb_SlipNo.Text & "';"
                End If
            Next

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()
            End If

            sqlTran.Commit()
        Catch ex As MySqlException
            sqlTran.Rollback()
            thread_LoadingFormEnd
            Thread.Sleep(100)
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_msgbox_string)
            Exit Sub
        End Try

        DBClose()

        thread_LoadingFormEnd
        Thread.Sleep(100)

        Load_ShipInfo(tb_SlipNo.Text)
        ReportLoad()
        MsgBox("출고 정보를 저장하였습니다.", MsgBoxStyle.Information, form_msgbox_string)

        tb_SlipNo.Text = String.Empty
        grid_ShipList.Rows.Count = 2
        btn_Save.Enabled = False
        btn_PDA_TempLoad.Enabled = False

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
            strSQL = "delete from module_ship_information;"

            sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
            sqlCmd_MDB.Transaction = sqlTran_MDB
            sqlCmd_MDB.ExecuteNonQuery()

            strSQL = "delete from module_ship_lot_information;"

            sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
            sqlCmd_MDB.Transaction = sqlTran_MDB
            sqlCmd_MDB.ExecuteNonQuery()

            strSQL = "delete from module_ship_part_information;"

            sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
            sqlCmd_MDB.Transaction = sqlTran_MDB
            sqlCmd_MDB.ExecuteNonQuery()

            sqlTran_MDB.Commit()

        Catch ex As OleDb.OleDbException
            sqlTran_MDB.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_msgbox_string)
            Exit Sub
        End Try

        '새로운 데이터를 MySQL에서 가져온다.
        sqlTran_MDB = mdbConnection1.BeginTransaction

        Try
            DBConnect()

            strSQL = "call sp_module_ship(2"
            strSQL += ", null"
            strSQL += ", null"
            strSQL += ",'" & slip_no & "'"
            strSQL += ",null)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                strSQL = "insert into module_ship_information(slip_no, ship_date, lot_no) values"
                strSQL += "('" & sqlDR("slip_no") & "'"
                strSQL += ",'" & sqlDR("ship_date") & "'"
                strSQL += ",'" & sqlDR("lot_no") & "');"

                sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
                sqlCmd_MDB.Transaction = sqlTran_MDB
                sqlCmd_MDB.ExecuteNonQuery()
            Loop
            sqlDR.Close()

            strSQL = "call sp_module_ship(3"
            strSQL += ", null"
            strSQL += ", null"
            strSQL += ",'" & slip_no & "'"
            strSQL += ",null)"

            sqlCmd = New MySqlCommand(strSQL, dbConnection1)
            sqlDR = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                strSQL = "insert into module_ship_lot_information(yj_no, product, lot_no"
                strSQL += ", chip_qty, ship_note) values"
                strSQL += "('" & sqlDR("yj_no") & "'"
                strSQL += ",'" & sqlDR("product") & "'"
                strSQL += ",'" & sqlDR("lot_no") & "'"
                strSQL += ",'" & sqlDR("chip_qty") & "'"
                strSQL += ",'" & sqlDR("ship_note") & "');"

                sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
                sqlCmd_MDB.Transaction = sqlTran_MDB
                sqlCmd_MDB.ExecuteNonQuery()
            Loop
            sqlDR.Close()

            strSQL = "call sp_module_ship(4"
            strSQL += ", null"
            strSQL += ", null"
            strSQL += ",'" & slip_no & "'"
            strSQL += ",null)"

            sqlCmd = New MySqlCommand(strSQL, dbConnection1)
            sqlDR = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                strSQL = "insert into module_ship_part_information(part_division, part_no"
                strSQL += ", lot_no, used_qty, yj_no, pfq_doe) values"
                strSQL += "('" & sqlDR("part_division") & "'"
                strSQL += ",'" & sqlDR("part_no") & "'"
                strSQL += ",'" & sqlDR("lot_no") & "'"
                strSQL += ",'" & sqlDR("kitting_qty") & "'"
                strSQL += ",'" & sqlDR("yj_no") & "'"
                strSQL += ",'" & sqlDR("pfq_doe") & "');"

                sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
                sqlCmd_MDB.Transaction = sqlTran_MDB
                sqlCmd_MDB.ExecuteNonQuery()
            Loop
            sqlDR.Close()

            DBClose()

            sqlTran_MDB.Commit()

        Catch ex As OleDb.OleDbException
            sqlTran_MDB.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_msgbox_string)
            Exit Sub
        End Try

        MDBClose()

    End Sub

    Private Sub ReportLoad()

        '크리스탈 레포트에서 인쇄하는 규칙
        Try
            Dim rptPath As String = Application.StartupPath & "\Reports\rpt_Module_Ship.rpt"

            Dim ds As DataSet = New DataSet

            Dim connection As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath + "\TempDB\TempDB.mdb" & ";Jet OLEDB:Database Password='dbwlspark'")
            connection.Open()

            Dim strQuery As String = " SELECT `module_ship_lot_information`.`product`, `module_ship_lot_information`.`lot_no`, "
            strQuery += " `module_ship_lot_information`.`chip_qty`, `module_ship_information`.`slip_no`, "
            strQuery += " `module_ship_information`.`ship_date`, `module_ship_lot_information`.`ship_note`"
            strQuery += " FROM   `module_ship_lot_information` `module_ship_lot_information` "
            strQuery += " INNER JOIN `module_ship_information` `module_ship_information` "
            strQuery += " ON `module_ship_lot_information`.`lot_no`=`module_ship_information`.`lot_no`"

            'Console.WriteLine(strQuery)

            Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(strQuery, connection)
            da.Fill(ds)

            Dim rDOC As ReportDocument = New ReportDocument

            rDOC.Load(rptPath)
            rDOC.SetDataSource(ds)
            'rDOC.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape '용지방향 설정
            frm_Report_Print.CrystalReportViewer1.ReportSource = Nothing
            frm_Report_Print.CrystalReportViewer1.ReportSource = rDOC

            If MsgBox("인쇄 미리보기를 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.Yes Then
                frm_Report_Print.ShowDialog()
            Else
                rDOC.PrintToPrinter(2, True, 0, 0)
            End If

            connection.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_msgbox_string)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Load_ShipInfo(tb_SlipNo.Text)
        ReportLoad()

    End Sub

    Private Sub grid_OutList_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles grid_OutList.MouseDoubleClick

        If e.Button = MouseButtons.Left And grid_OutList.MouseRow > 0 Then
            Dim selRow As Integer = grid_OutList.MouseRow
            Dim slip_no As String = grid_OutList(selRow, 2)
            thread_LoadingFormStart()

            grid_status = "Editing"
            btn_Save.Enabled = True
            btn_TempLoad.Enabled = False
            btn_TempSave.Enabled = False

            Dim total_Module As Integer = 0

            grid_ShipList.Redraw = False
            grid_ShipList.Rows.Count = 2

            tb_SlipNo.Text = slip_no
            Load_LotList(slip_no)

            For i = 2 To grid_ShipList.Rows.Count - 1
                Load_lot_information(grid_ShipList(i, 3), i)
                If Not IsDBNull(grid_ShipList(i, 1)) Then
                    Load_Repair_Data(grid_ShipList(i, 1), i, True)
                End If
                total_Module += grid_ShipList(i, 4)
            Next

            grid_ShipList.AutoSizeCols()
            grid_ShipList.AutoSizeRows(2, 0, grid_ShipList.Rows.Count - 1, grid_ShipList.Cols.Count - 1, 0, AutoSizeFlags.None)
            grid_ShipList.Redraw = True

            Label10.Text = "총 출고 Lot수 : " & Format(grid_ShipList.Rows.Count - 2, "#,##0") & " Lot"
            Label3.Text = "총 출고 Module : " & Format(total_Module, "#,##0 EA")

            thread_LoadingFormEnd
        End If

    End Sub

    Private Sub Load_LotList(ByVal slip_no As String)

        DBConnect()

        Dim strSQL As String = "select lot_no" &
            " from module_ship_information" &
            " where slip_no = '" & slip_no & "'" &
            " order by ship_note"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_ShipList.Rows.Count - 1 & vbTab &
                                          vbTab &
                                          vbTab &
                                          sqlDR("lot_no")
            grid_ShipList.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

    End Sub

    Private Sub Load_lot_information(ByVal lot_no As String, ByVal editRow As Integer)

        DBConnect()

        Dim strSql As String = "call sp_module_ship(7" &
            ", null" &
            ", null" &
            ", '" & lot_no & "'" &
            ", null)"
        Dim sqlCmd As New MySqlCommand(strSql, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            grid_ShipList(editRow, 1) = sqlDR("yj_no")
            grid_ShipList(editRow, 2) = sqlDR("product")
            grid_ShipList(editRow, 4) = sqlDR("chip_qty")
            grid_ShipList(editRow, 8) = sqlDR("ship_note")
        Loop

        sqlDR.Close()
        DBClose()

    End Sub

    Private Sub grid_ShipList_MouseMove(sender As Object, e As MouseEventArgs) Handles grid_ShipList.MouseMove, grid_OutList.MouseMove

        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then

            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub

    Private Sub grid_LotList_LostFocus(sender As Object, e As EventArgs) Handles grid_ShipList.LostFocus,
                                                                                 grid_OutList.LostFocus

        frm_Main.lb_Status.Text = String.Empty

    End Sub

    Private Sub tb_SlipNo_TextChanged(sender As Object, e As EventArgs) Handles tb_SlipNo.TextChanged

    End Sub

    Private Sub btn_TempSave_Click(sender As Object, e As EventArgs) Handles btn_TempSave.Click

        If MsgBox("임시 저장 하시겠습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                  form_msgbox_string) = MsgBoxResult.No Then Exit Sub

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try
            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            '라벨 발행 기록
            For i = 1 To grid_ShipList.Rows.Count - 1
                If grid_ShipList(i, 0) = "N" Then
                    strSQL += "insert into module_ship_information_temp(yj_no, ship_note) values("
                    strSQL += "'" & grid_ShipList(i, 1) & "'"
                    strSQL += ",'" & grid_ShipList(i, 8) & "');"
                ElseIf grid_ShipList(i, 0) = "M" Then
                    strSQL += "update module_ship_information_temp set ship_note = "
                    strSQL += "'" & grid_ShipList(i, 8) & "'"
                    strSQL += " where yj_no = '" & grid_ShipList(i, 1) & "';"
                ElseIf grid_ShipList(i, 0) = "D" Then
                    strSQL += "delete from module_ship_information_temp"
                    strSQL += " where yj_no = '" & grid_ShipList(i, 1) & "';"
                End If
            Next

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()
            End If

            sqlTran.Commit()
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_msgbox_string)
            Exit Sub
        End Try

        DBClose()

        MsgBox("임시 출고 정보를 저장하였습니다.", MsgBoxStyle.Information, form_msgbox_string)

    End Sub

    Private Sub btn_TempLoad_Click(sender As Object, e As EventArgs) Handles btn_TempLoad.Click

        Dim msgString As String = "임시데이터를 불러온 후 서버에서 내용이 삭제 됩니다." & vbCrLf &
            "다시 데이터를 활용하려면 임시저장을 다시 해주십시오."

        MsgBox(msgString, MsgBoxStyle.Information, form_msgbox_string)

        grid_ShipList.Redraw = False
        grid_ShipList.Rows.Count = 2

        DBConnect()

        Dim strSQL As String = "select yj_no, ship_note from module_ship_information_temp"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = "N" & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          vbTab &
                                          vbTab &
                                          vbTab &
                                          vbTab &
                                          vbTab &
                                          vbTab &
                                          sqlDR("ship_note")
            grid_ShipList.AddItem(insert_String)
            grid_ShipList.Rows(grid_ShipList.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
        Loop
        sqlDR.Close()

        Dim sqlTran As MySqlTransaction

        sqlTran = dbConnection1.BeginTransaction

        Try
            strSQL = "delete from module_ship_information_temp;"

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()
            End If

            sqlTran.Commit()
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_msgbox_string)
            Exit Sub
        End Try

        DBClose()

        For i = 2 To grid_ShipList.Rows.Count - 1
            Load_Lot_Status(grid_ShipList(i, 1), i)
            Load_Repair_Data(grid_ShipList(i, 1), i, False)
        Next

        grid_ShipList.AutoSizeCols()
        grid_ShipList.AutoSizeRows(2, 0, grid_ShipList.Rows.Count - 1, grid_ShipList.Cols.Count - 1, 0, AutoSizeFlags.None)
        grid_ShipList.Redraw = True

        Label10.Text = "총 출고 Lot수 : " & Format(grid_ShipList.Rows.Count - 2, "#,##0") & " Lot"

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub btn_PDA_Check_Click(sender As Object, e As EventArgs) Handles btn_PDA_Check.Click

        If Not frm_Module_Out_PDA.Visible Then frm_Module_Out_PDA.Show()
        frm_Module_Out_PDA.Focus()

    End Sub

    Private Sub btn_PDA_TempLoad_Click(sender As Object, e As EventArgs) Handles btn_PDA_TempLoad.Click

        Dim msgString As String = "출력물과 실물비교 완료 하였습니까?"

        If MsgBox(msgString,
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                  form_msgbox_string) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart()

        grid_ShipList.Redraw = False
        grid_ShipList.Rows.Count = 2
        Dim total_Module As Integer = 0

        DBConnect()

        Dim strSQL As String = "call sp_app_module_ship_temp(0)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = "N" & vbTab &
                                          String.Empty & vbTab &
                                          String.Empty & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          String.Empty & vbTab &
                                          String.Empty & vbTab &
                                          String.Empty & vbTab &
                                          String.Empty & vbTab &
                                          String.Empty
            grid_ShipList.AddItem(insert_String)
            grid_ShipList.Rows(grid_ShipList.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
        Loop
        sqlDR.Close()
        DBClose()

        For i = 2 To grid_ShipList.Rows.Count - 1
            Load_Lot_Status(grid_ShipList(i, 3), i)
            Load_Repair_Data(grid_ShipList(i, 3), i, False)
            total_Module += grid_ShipList(i, 4)
        Next

        grid_ShipList.AutoSizeCols()
        grid_ShipList.AutoSizeRows(2, 0, grid_ShipList.Rows.Count - 1, grid_ShipList.Cols.Count - 1, 0, AutoSizeFlags.None)
        grid_ShipList.Redraw = True

        Label10.Text = "총 출고 Lot수 : " & Format(grid_ShipList.Rows.Count - 2, "#,##0") & " Lot"
        Label3.Text = "총 출고 Module : " & Format(total_Module, "#,##0 EA")

        thread_LoadingFormEnd

    End Sub

    Private Sub grid_OutList_KeyDown(sender As Object, e As KeyEventArgs) Handles grid_OutList.KeyDown

        If e.Shift Then
            md_ETC.cellCal(sender, Nothing)
        Else
            frm_Main.lb_Status.Text = String.Empty
        End If

    End Sub
End Class
Imports System.Threading
Imports C1.Win.C1FlexGrid
Imports CrystalDecisions.CrystalReports.Engine
Imports MySql.Data.MySqlClient

Public Class frm_WorkSite_Moving

    Dim form_Msgbox_String As String = "현장출고 등록"

    Private Sub frm_WorkSite_Moving_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        GridSetting()

        DateTimePicker1.Value = Format(Now, "yyyy-MM-01 00:00:00")
        DateTimePicker2.Value = Format(Now, "yyyy-MM-dd 23:59:59")

    End Sub

    Private Sub GridSetting()

        With grid_KittingList
            .AllowEditing = False
            .AllowFiltering = True
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 8
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_KittingList(0, 0) = "No"
            grid_KittingList(0, 1) = "선택"
            grid_KittingList(0, 2) = "Product"
            grid_KittingList(0, 3) = "Lot No."
            grid_KittingList(0, 4) = "YJ No."
            grid_KittingList(0, 5) = "Module Qty"
            grid_KittingList(0, 6) = "수리구분"
            grid_KittingList(0, 7) = "Baking 일자"
            .Cols(1).DataType = GetType(Boolean)
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

        With grid_SelectList
            .AllowEditing = True
            .AllowFiltering = True
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 9
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_SelectList(0, 0) = "No"
            grid_SelectList(0, 1) = "Product"
            grid_SelectList(0, 2) = "Lot No."
            grid_SelectList(0, 3) = "YJ No."
            grid_SelectList(0, 4) = "Module Qty"
            grid_SelectList(0, 5) = "수리구분"
            grid_SelectList(0, 6) = "Baking 일자"
            grid_SelectList(0, 7) = "비고"
            grid_SelectList(0, 8) = "temp code"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 2).TextAlign = TextAlignEnum.LeftCenter
            .Cols(8).Visible = False
            .AutoSizeCols()
            .ShowCursor = True
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
            grid_OutList(0, 1) = "전표번호"
            grid_OutList(0, 2) = "출고일시"
            grid_OutList(0, 3) = "출고 Lot수"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 2).TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
            '.ExtendLastCol = True
        End With

    End Sub

    Private Sub btn_New_Kitting_Click(sender As Object, e As EventArgs) Handles btn_New_Kitting.Click

        thread_LoadingFormStart()

        tb_KONo.Text = "KO" & Format(Now, "yyyyMMddHHmmss")

        grid_SelectList.Rows.Count = 1
        grid_KittingList.Rows.Count = 1
        grid_KittingList.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_worksite_moving(0" &
            ", null" &
            ", null" &
            ", null" &
            ", null" &
            ", null)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_string As String = grid_KittingList.Rows.Count & vbTab &
                                          vbTab &
                                          sqlDR("product") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          Format(sqlDR("chip_qty"), "#,##0") & vbTab &
                                          sqlDR("repair_section") & vbTab &
                                          Format(CDate(sqlDR("kitting_date")), "yyyy-MM-dd HH:mm:ss")
            grid_KittingList.AddItem(insert_string)
        Loop
        sqlDR.Close()

        DBClose()

        grid_KittingList.AutoSizeCols()
        grid_KittingList.Redraw = True

        btn_Save.Enabled = True

        thread_LoadingFormEnd

    End Sub

    Private Sub grid_KittingList_Click(sender As Object, e As EventArgs) Handles grid_KittingList.Click

    End Sub

    Private Sub grid_KittingList_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles grid_KittingList.MouseDoubleClick

        If e.Button = MouseButtons.Left And grid_KittingList.MouseRow > 0 Then

            grid_SelectList.Redraw = False

            Dim selRow As Integer = grid_KittingList.MouseRow
            If grid_KittingList.GetCellCheck(selRow, 1) = CheckEnum.Checked Then
                grid_KittingList.SetCellCheck(selRow, 1, CheckEnum.Unchecked)
                Dim sel_lot_No As String = grid_KittingList(selRow, 3)
                Dim findRow As Integer = grid_SelectList.FindRow(sel_lot_No, 1, 2, False)
                grid_SelectList.RemoveItem(findRow)
            Else
                grid_KittingList.SetCellCheck(selRow, 1, CheckEnum.Checked)
                Dim insertString As String = grid_SelectList.Rows.Count & vbTab &
                    grid_KittingList(selRow, 2) & vbTab &
                    grid_KittingList(selRow, 3) & vbTab &
                    grid_KittingList(selRow, 4) & vbTab &
                    grid_KittingList(selRow, 5) & vbTab &
                    grid_KittingList(selRow, 6) & vbTab &
                    grid_KittingList(selRow, 7)
                grid_SelectList.AddItem(insertString)
                grid_SelectList.AutoSizeCols()
            End If

            For i = 1 To grid_SelectList.Rows.Count - 1
                grid_SelectList(i, 0) = i
            Next

            Label12.Text = grid_SelectList.Rows.Count - 1
            grid_SelectList.Redraw = True
        End If

    End Sub

    Private Sub grid_SelectList_RowColChange(sender As Object, e As EventArgs) Handles grid_SelectList.RowColChange

        If Not grid_SelectList.Col = grid_SelectList.Cols.Count - 1 Then
            grid_SelectList.AllowEditing = False
        Else
            grid_SelectList.AllowEditing = True
        End If

    End Sub

    Private Sub btn_Save_Click(sender As Object, e As EventArgs) Handles btn_Save.Click

        If Trim(tb_Person1.Text) = String.Empty Or
            Trim(tb_Person2.Text) = String.Empty Or
            Trim(tb_Person3.Text) = String.Empty Then
            MsgBox("인계자, 인수자, 확인자가 모두 입력되지 않았습니다.",
                    MsgBoxStyle.Information,
                    form_Msgbox_String)
            Exit Sub
        End If

        If grid_SelectList.Rows.Count = 1 Then
            MsgBox("출고 항목이 선택되지 않았습니다.",
                   MsgBoxStyle.Information,
                   form_Msgbox_String)
            Exit Sub
        End If

        thread_LoadingFormStart("Saving...")

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try
            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            strSQL = "insert kitting_out_information(ko_no, ko_person1, ko_person2, ko_person3, ko_date) values("
            strSQL += "'" & tb_KONo.Text & "'"
            strSQL += ", '" & tb_Person1.Text & "'"
            strSQL += ", '" & tb_Person2.Text & "'"
            strSQL += ", '" & tb_Person3.Text & "'"
            strSQL += ", '" & write_date & "');"

            For i = 1 To grid_SelectList.Rows.Count - 1
                strSQL += "insert kitting_out_information_detail(temp_code, ko_no, lot_no, yj_no, ko_note"
                strSQL += ", writer, write_date, first_write_date) values("
                strSQL += "'" & tb_KONo.Text & i & "'"
                strSQL += ",'" & tb_KONo.Text & "'"
                strSQL += ",'" & grid_SelectList(i, 2) & "'"
                strSQL += ",'" & grid_SelectList(i, 3) & "'"
                strSQL += ",'" & grid_SelectList(i, 7) & "'"
                strSQL += ",'" & loginID & "'"
                strSQL += ",'" & write_date & "'"
                strSQL += ",'" & write_date & "');"

                strSQL += "update basic_lot_information set lot_status = 'Moving to WorkSite'"
                strSQL += " where lot_no = '" & grid_SelectList(i, 2) & "';"
            Next

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

        MovingList_ServerLoad()
        ReportLoad()

        btn_Save.Enabled = False
        tb_KONo.Text = String.Empty
        tb_KO_Date.Text = String.Empty

        grid_KittingList.Rows.Count = 1
        grid_SelectList.Rows.Count = 1

        thread_LoadingFormEnd
        Thread.Sleep(100)

        MsgBox("현장 출고 등록을 완료 하였습니다.",
               MsgBoxStyle.Information,
               form_Msgbox_String)

    End Sub

    Private Sub MovingList_ServerLoad()

        Dim strSQL As String = String.Empty

        'MySQL DB에서 정보를 불러오기전에 기존 내용 삭제
        Mdbconnect()

        Dim sqlTran_MDB As OleDb.OleDbTransaction
        Dim sqlCmd_MDB As OleDb.OleDbCommand

        sqlTran_MDB = mdbConnection1.BeginTransaction

        Try
            '기존 저장 데이터를 삭제
            strSQL = "delete from kitting_out_information;"

            sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
            sqlCmd_MDB.Transaction = sqlTran_MDB
            sqlCmd_MDB.ExecuteNonQuery()

            strSQL = "delete from kitting_out_information_detail;"

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

            strSQL = "call sp_worksite_moving(1"
            strSQL += ", null"
            strSQL += ", null"
            strSQL += ",'" & tb_KONo.Text & "'"
            strSQL += ",null"
            strSQL += ",null)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                strSQL = "insert into kitting_out_information(ko_no, ko_person1, ko_person2"
                strSQL += ", ko_person3, ko_date) values"
                strSQL += "('" & sqlDR("ko_no") & "'"
                strSQL += ",'" & sqlDR("ko_person1") & "'"
                strSQL += ",'" & sqlDR("ko_person2") & "'"
                strSQL += ",'" & sqlDR("ko_person3") & "'"
                strSQL += ",'" & sqlDR("ko_date") & "');"

                sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, mdbConnection1)
                sqlCmd_MDB.Transaction = sqlTran_MDB
                sqlCmd_MDB.ExecuteNonQuery()
            Loop
            sqlDR.Close()

            strSQL = "call sp_worksite_moving(2"
            strSQL += ", null"
            strSQL += ", null"
            strSQL += ",'" & tb_KONo.Text & "'"
            strSQL += ",null"
            strSQL += ",null)"

            sqlCmd = New MySqlCommand(strSQL, dbConnection1)
            sqlDR = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                strSQL = "insert into kitting_out_information_detail(temp_code, ko_no, lot_no, yj_no"
                strSQL += ", ko_note, writer, write_date, first_write_date, product, chip_qty, repair_section) values"
                strSQL += "('" & sqlDR("temp_code") & "'"
                strSQL += ",'" & sqlDR("ko_no") & "'"
                strSQL += ",'" & sqlDR("lot_no") & "'"
                strSQL += ",'" & sqlDR("yj_no") & "'"
                strSQL += ",'" & sqlDR("ko_note") & "'"
                strSQL += ",'" & sqlDR("writer") & "'"
                strSQL += ",'" & sqlDR("write_date") & "'"
                strSQL += ",'" & sqlDR("first_write_date") & "'"
                strSQL += ",'" & sqlDR("product") & "'"
                strSQL += ",'" & sqlDR("chip_qty") & "'"
                strSQL += ",'" & sqlDR("repair_section") & "');"

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
            Dim rptPath As String = Application.StartupPath & "\Reports\WorkSite Moving.rpt"

            Dim ds As DataSet = New DataSet

            Dim connection As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath + "\TempDB\TempDB.mdb" & ";Jet OLEDB:Database Password='dbwlspark'")
            connection.Open()

            Dim strQuery As String = "SELECT `kitting_out_information_detail`.`temp_code`, `kitting_out_information_detail`.`lot_no`,"
            strQuery += "`kitting_out_information_detail`.`yj_no`, `kitting_out_information_detail`.`product`, `kitting_out_information_detail`.`chip_qty`,"
            strQuery += "`kitting_out_information_detail`.`repair_section`, `kitting_out_information`.`ko_no`, `kitting_out_information`.`ko_person1`,"
            strQuery += "`kitting_out_information`.`ko_person2`, `kitting_out_information`.`ko_person3`, `kitting_out_information`.`ko_date`,"
            strQuery += "`kitting_out_information_detail`.`ko_note`"
            strQuery += " FROM   `kitting_out_information_detail` `kitting_out_information_detail`"
            strQuery += " INNER JOIN `kitting_out_information` `kitting_out_information`"
            strQuery += " ON `kitting_out_information_detail`.`ko_no`=`kitting_out_information`.`ko_no`"

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
                rDOC.PrintToPrinter(1, True, 0, 0)
            End If

            connection.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
        End Try

    End Sub

    Private Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        thread_LoadingFormStart()

        btn_Save.Enabled = False
        grid_KittingList.Rows.Count = 1
        grid_SelectList.Rows.Count = 1
        tb_KONo.Text = String.Empty
        tb_KO_Date.Text = String.Empty

        grid_OutList.Rows.Count = 1
        grid_OutList.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_worksite_moving(3" &
            ", '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "'" &
            ", '" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker2.Value), "yyyy-MM-dd") & "'" &
            ", '" & TB_S_YJ_No.Text & "'" &
            ", null" &
            ", null)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_string As String = grid_OutList.Rows.Count & vbTab &
                                          sqlDR("ko_no") & vbTab &
                                          Format(sqlDR("ko_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          Format(sqlDR("ship_lot_count"), "#,##0")
            grid_OutList.AddItem(insert_string)
        Loop
        sqlDR.Close()

        DBClose()

        grid_OutList.AutoSizeCols()
        grid_OutList.Redraw = True

        thread_LoadingFormEnd

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub grid_OutList_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles grid_OutList.MouseDoubleClick

        If e.Button = MouseButtons.Left And grid_OutList.MouseRow > 0 Then
            Dim selrow As Integer = grid_OutList.MouseRow
            thread_LoadingFormStart()
            Ship_Info_Load(grid_OutList(selrow, 1))
            Ship_Info_Detail_Load(grid_OutList(selrow, 1))
            grid_KittingList.Rows.Count = 1
            thread_LoadingFormEnd
        End If

    End Sub

    Private Sub Ship_Info_Load(ByVal sel_KO_No As String)

        DBConnect()

        Dim strSQL As String = "call sp_worksite_moving(1" &
            ", null" &
            ", null" &
            ", '" & sel_KO_No & "'" &
            ", null" &
            ", null)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            tb_KONo.Text = sqlDR("ko_no")
            tb_KO_Date.Text = Format(sqlDR("ko_date"), "yyyy-MM-dd HH:mm:ss")
            tb_Person1.Text = sqlDR("ko_person1")
            tb_Person2.Text = sqlDR("ko_person2")
            tb_Person3.Text = sqlDR("ko_person3")
        Loop
        sqlDR.Close()

        DBClose()

    End Sub

    Private Sub Ship_Info_Detail_Load(ByVal sel_KO_No As String)

        grid_SelectList.Rows.Count = 1
        grid_SelectList.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_worksite_moving(2" &
            ", null" &
            ", null" &
            ", '" & sel_KO_No & "'" &
            ", null" &
            ", null)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_string As String = grid_SelectList.Rows.Count & vbTab &
                                          sqlDR("product") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          Format(sqlDR("chip_qty"), "#,##0") & vbTab &
                                          sqlDR("repair_section") & vbTab &
                                          Format(CDate(sqlDR("kitting_date")), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          sqlDR("ko_note") & vbTab &
                                          sqlDR("temp_code")
            grid_SelectList.AddItem(insert_string)
        Loop
        sqlDR.Close()

        DBClose()

        grid_SelectList.AutoSizeCols()
        grid_SelectList.Redraw = True
        Label12.Text = grid_SelectList.Rows.Count - 1

    End Sub

    Private Sub grid_OutList_MouseEnterCell(sender As Object, e As RowColEventArgs) Handles grid_OutList.MouseEnterCell

    End Sub

    Private Sub grid_KittingList_MouseMove(sender As Object, e As MouseEventArgs) Handles grid_KittingList.MouseMove,
                                                                                          grid_OutList.MouseMove,
                                                                                          grid_SelectList.MouseMove

        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then

            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub

    Private Sub grid_LotList_LostFocus(sender As Object, e As EventArgs) Handles grid_KittingList.LostFocus,
                                                                                 grid_OutList.LostFocus,
                                                                                 grid_SelectList.LostFocus

        frm_Main.lb_Status.Text = String.Empty

    End Sub

    Private Sub grid_SelectList_KeyDown(sender As Object, e As KeyEventArgs) Handles grid_SelectList.KeyDown,
                                                                                     grid_KittingList.KeyDown,
                                                                                     grid_OutList.KeyDown

        If e.Shift Then
            md_ETC.cellCal(sender, Nothing)
        Else
            frm_Main.lb_Status.Text = String.Empty
        End If

    End Sub
End Class
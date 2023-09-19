'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-09-11
'##### 수정일자 : 2023-09-11
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 

'############################################################################################################
'############################################################################################################

Imports C1.Win.C1FlexGrid
Imports CrystalDecisions.CrystalReports.Engine
Imports MySql.Data.MySqlClient

Public Class frm_Module_Out_PDA

    Dim form_Msgbox_String As String = "PDA 등록 리스트 확인"

    Private Sub frm_Module_Out_PDA_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        gridSetting()

        BTN_Search_Click(Nothing, Nothing)

    End Sub

    Private Sub gridSetting()

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
            grid_OutList(0, 1) = "YJ No."
            grid_OutList(0, 2) = "Lot No."
            grid_OutList(0, 3) = "Module Qty"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

    End Sub

    Private Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        thread_LoadingFormStart()

        btn_Print.Enabled = True
        grid_OutList.Redraw = False
        grid_OutList.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call sp_app_module_ship_temp(0)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_OutList.Rows.Count & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          Format(sqlDR("chip_qty"), "#,##0")
            grid_OutList.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        grid_OutList.AutoSizeCols()
        grid_OutList.Redraw = True

        th_LoadingWindow.Abort()

    End Sub

    Private Sub btn_Print_Click(sender As Object, e As EventArgs) Handles btn_Print.Click

        listLoad()
        ReportLoad()

    End Sub

    Private Sub listLoad()

        Dim strSQL As String = String.Empty

        'MySQL DB에서 정보를 불러오기전에 기존 내용 삭제
        Mdbconnect()

        Dim sqlTran_MDB As OleDb.OleDbTransaction
        Dim sqlCmd_MDB As OleDb.OleDbCommand

        sqlTran_MDB = mdbConnection1.BeginTransaction

        Try
            '기존 저장 데이터를 삭제
            strSQL = "delete from module_ship_temp;"

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

            strSQL = "call sp_app_module_ship_temp(0)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                strSQL = "insert into module_ship_temp(yj_no, lot_no, chip_qty) values"
                strSQL += "('" & sqlDR("yj_no") & "'"
                strSQL += ",'" & sqlDR("lot_no") & "'"
                strSQL += ",'" & sqlDR("chip_qty") & "');"

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
            Dim rptPath As String = Application.StartupPath & "\Reports\rpt_Module_Ship_Check.rpt"

            Dim ds As DataSet = New DataSet

            Dim connection As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath + "\TempDB\TempDB.mdb" & ";Jet OLEDB:Database Password='dbwlspark'")
            connection.Open()

            Dim strQuery As String = " SELECT `module_ship_temp`.`yj_no`, `module_ship_temp`.`lot_no`, `module_ship_temp`.`chip_qty`"
            strQuery += " FROM   `module_ship_temp` `module_ship_temp`"
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

    Private Sub grid_OutList_MouseClick(sender As Object, e As MouseEventArgs) Handles grid_OutList.MouseClick

        If e.Button = MouseButtons.Right And grid_OutList.MouseRow > 0 Then
            cms_Menu.Show(grid_OutList, New Point(e.X, e.Y))
        End If

    End Sub

    Private Sub btn_RowDelete_Click(sender As Object, e As EventArgs) Handles btn_RowDelete.Click


        If grid_OutList(grid_OutList.Row, 0).ToString = "N" Then
            grid_OutList.RemoveItem(grid_OutList.Row)
        Else
            grid_OutList(grid_OutList.Row, 0) = "D"
        End If

    End Sub

    Private Sub btn_Save2_Click(sender As Object, e As EventArgs) Handles btn_Save2.Click

        If MsgBox("저장 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try

            For i = 1 To grid_OutList.Rows.Count - 1
                If grid_OutList(i, 0).ToString = "D" Then
                    strSQL += "delete from module_ship_information_pda_temp"
                    strSQL += " where lot_no = '" & grid_OutList(i, 2) & "';"
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

        MsgBox("저장완료.", MsgBoxStyle.Information, form_Msgbox_String)

        BTN_Search_Click(Nothing, Nothing)

    End Sub
End Class
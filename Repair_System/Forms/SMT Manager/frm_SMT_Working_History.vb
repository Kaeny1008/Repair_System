Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_SMT_Working_History

    Dim form_Msgbox_String As String = "SMT 생산현황"

    Private Sub frm_SMT_Working_History_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        NumericUpDown1.Value = 10
        timer_SMT_Working.Interval = 10000
        timer_SMT_Working.Enabled = True
        Timer1.Interval = 20
        Timer1.Enabled = False
        Timer2.Interval = 500
        Timer2.Enabled = False

        Start_Date.Value = Format(Now, "yyyy-MM-01 00:00:00")
        End_Date.Value = Format(Now, "yyyy-MM-dd 23:59:59")

        GRID_SETTING()

        load_SMT_Working()

    End Sub

    Private Sub GRID_SETTING()

        With grid_SMT_Working_Data
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
            grid_SMT_Working_Data(0, 0) = "No"
            grid_SMT_Working_Data(0, 1) = "Part No."
            grid_SMT_Working_Data(0, 2) = "Lot No."
            grid_SMT_Working_Data(0, 3) = "사용수량"
            grid_SMT_Working_Data(0, 4) = "Loss수량"
            grid_SMT_Working_Data(0, 5) = "투입일자"
            grid_SMT_Working_Data(0, 6) = "종료일자"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
        End With

        With grid_Daily_List
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .AllowMerging = AllowMergingEnum.RestrictRows
            .Cols(1).AllowMerging = True
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 5
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_Daily_List(0, 0) = "No"
            grid_Daily_List(0, 1) = "생산일자"
            grid_Daily_List(0, 2) = "작업구분"
            grid_Daily_List(0, 3) = "Lot 수"
            grid_Daily_List(0, 4) = "Module 수"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
            .SelectionMode = SelectionModeEnum.Row
        End With

        With C1FlexGrid1
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 12
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            C1FlexGrid1(0, 0) = "No"
            C1FlexGrid1(0, 1) = "Product Code"
            C1FlexGrid1(0, 2) = "Bucket"
            C1FlexGrid1(0, 3) = "Lot No."
            C1FlexGrid1(0, 4) = "Option"
            C1FlexGrid1(0, 5) = "YJ No."
            C1FlexGrid1(0, 6) = "Module Qty"
            C1FlexGrid1(0, 7) = "작업구분"
            C1FlexGrid1(0, 8) = "시작일자"
            C1FlexGrid1(0, 9) = "종료일자"
            C1FlexGrid1(0, 10) = "작업시간(분)"
            C1FlexGrid1(0, 11) = "상태"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(11).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
            .ExtendLastCol = True
            .Cols(6).Format = "#,##0"
            .Cols(10).Format = "#,##0"
        End With

        'set up styles
        Dim s As CellStyle = C1FlexGrid1.Styles(CellStyleEnum.Subtotal0)
        s.BackColor = Color.DarkBlue
        s.ForeColor = Color.White
        s.Font = New Font(C1FlexGrid1.Font, FontStyle.Bold)

        s = C1FlexGrid1.Styles(CellStyleEnum.Subtotal1)
        s.BackColor = Color.LightBlue
        s.ForeColor = Color.Black
        s = C1FlexGrid1.Styles(CellStyleEnum.Subtotal2)
        s.BackColor = Color.DarkRed
        s.ForeColor = Color.White

        'more setup
        C1FlexGrid1.AllowDragging = AllowDraggingEnum.None
        C1FlexGrid1.AllowEditing = False
        'C1FlexGrid1.Cols(0).WidthDisplay \= 3
        C1FlexGrid1.Tree.Column = 1
        C1FlexGrid1.SubtotalPosition = SubtotalPositionEnum.AboveData

    End Sub

    Private Sub load_SMT_Working()

        tb_Slip_No.Text = String.Empty
        TextBox1.Text = String.Empty
        TextBox2.Text = String.Empty
        TextBox3.Text = String.Empty

        DBConnect()

        Dim strSQL As String = "call sp_smt_manager(0" &
                                ", null" &
                                ", null" &
                                ", null" &
                                ", null" &
                                ")"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            tb_Slip_No.Text = sqlDR("yj_no")
            TextBox1.Text = sqlDR("lot_no")
            TextBox2.Text = sqlDR("work_section")
            TextBox3.Text = Format(sqlDR("start_date"), "yyyy-MM-dd HH:mm:ss")
        Loop
        sqlDR.Close()

        DBClose()

        If tb_Slip_No.Text = String.Empty Then
            grid_SMT_Working_Data.Rows.Count = 1
            Label4.Visible = False
            Timer1.Enabled = False
            Timer2.Enabled = False
            Label9.Visible = False
        Else
            Label4.Visible = True
            Timer1.Enabled = True
            Label9.Visible = True
            Timer2.Enabled = True
            load_SMT_Working_Data()
        End If

    End Sub

    Private Sub load_SMT_Working_Data()

        grid_SMT_Working_Data.Rows.Count = 1
        grid_SMT_Working_Data.Redraw = False

        DBConnect()

        Dim strSQL As String = "call sp_smt_manager(1" &
                                ", null" &
                                ", null" &
                                ", '" & TextBox1.Text & "'" &
                                ", '" & TextBox2.Text & "'" &
                                ")"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insertString As String = grid_SMT_Working_Data.Rows.Count & vbTab &
                sqlDR("material_part_no") & vbTab &
                sqlDR("material_lot_no") & vbTab &
                sqlDR("material_used_qty") & vbTab &
                sqlDR("material_loss_qty") & vbTab &
                Format(sqlDR("start_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                sqlDR("end_date")
            grid_SMT_Working_Data.AddItem(insertString)
        Loop
        sqlDR.Close()

        DBClose()

        grid_SMT_Working_Data.AutoSizeCols()
        grid_SMT_Working_Data.Redraw = True

    End Sub

    Private Sub timer_SMT_Working_Tick(sender As Object, e As EventArgs) Handles timer_SMT_Working.Tick

        load_SMT_Working()

    End Sub

    Private Sub btn_TimerReset_Click(sender As Object, e As EventArgs) Handles btn_TimerReset.Click

        If NumericUpDown1.Value = 0 Then
            timer_SMT_Working.Enabled = False
            tb_Slip_No.Text = String.Empty
            TextBox1.Text = String.Empty
            TextBox2.Text = String.Empty
            TextBox3.Text = String.Empty
            grid_SMT_Working_Data.Rows.Count = 1
            Label4.Visible = False
            Timer1.Enabled = False
            Timer2.Enabled = False
            Label9.Visible = False
        Else
            timer_SMT_Working.Interval = NumericUpDown1.Value * 1000
            timer_SMT_Working.Enabled = False
            timer_SMT_Working.Enabled = True
            load_SMT_Working()
        End If

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        'If Label9.Text = ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" Then
        '    Label9.Text = ">"
        'Else
        '    Label9.Text += ">"
        'End If

        If Label9.Text = "                                                              ▷" Then
            Label9.Text = "▷"
        Else
            Label9.Text = " " & Label9.Text
        End If

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        If Label4.Visible Then
            Label4.Visible = False
        Else
            Label4.Visible = True
        End If

    End Sub

    Private Sub btn_Search_Click(sender As Object, e As EventArgs) Handles btn_Search.Click

        thread_LoadingFormStart()

        grid_Daily_List.Rows.Count = 1
        grid_Daily_List.Redraw = False
        Dim color1 As Color = Color.LightBlue
        Dim color2 As Color = Color.White
        Dim pointColor As Color = Nothing

        DBConnect()

        Dim strSQL As String = "call sp_smt_manager(2" &
                                ", '" & Format(Start_Date.Value, "yyyy-MM-dd 00:00:00") & "'" &
                                ", '" & Format(End_Date.Value, "yyyy-MM-dd 23:59:59") & "'" &
                                ", null" &
                                ", null" &
                                ")"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insertString As String = grid_Daily_List.Rows.Count & vbTab &
                sqlDR("working_date") & vbTab &
                sqlDR("work_section") & vbTab &
                Format(sqlDR("lot_qty"), "#,##0") & vbTab &
                Format(sqlDR("module_qty"), "#,##0")
            If Not grid_Daily_List(grid_Daily_List.Rows.Count - 1, 1) = sqlDR("working_date") Then
                If pointColor = color2 Then
                    pointColor = color1
                Else
                    pointColor = color2
                End If
            End If
            grid_Daily_List.AddItem(insertString)
            grid_Daily_List.Rows(grid_Daily_List.Rows.Count - 1).StyleNew.BackColor = pointColor
        Loop
        sqlDR.Close()

        DBClose()

        grid_Daily_List.AutoSizeCols()
        grid_Daily_List.Redraw = True

        thread_LoadingFormEnd

    End Sub

    Private Sub grid_Daily_List_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles grid_Daily_List.MouseDoubleClick

        Dim gridRow As Integer = grid_Daily_List.MouseRow

        If e.Button = MouseButtons.Left And gridRow > 0 Then

            thread_LoadingFormStart()

            Dim selDate As String = grid_Daily_List(gridRow, 1)
            Dim selWorkSection As String = grid_Daily_List(gridRow, 2)

            C1FlexGrid1.Rows.Count = 1
            C1FlexGrid1.Redraw = False

            DBConnect()

            Dim strSQL As String = "call sp_smt_manager(3" &
                                    ", null" &
                                    ", null" &
                                    ", '" & selDate & "'" &
                                    ", '" & selWorkSection & "'" &
                                    ")"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                Dim insertString As String = C1FlexGrid1.Rows.Count & vbTab &
                    sqlDR("product") & vbTab &
                    sqlDR("bucket") & vbTab &
                    sqlDR("lot_no") & vbTab &
                    sqlDR("lot_option") & vbTab &
                    sqlDR("yj_no") & vbTab &
                    Format(sqlDR("chip_qty"), "#,##0") & vbTab &
                    sqlDR("work_section") & vbTab &
                    Format(sqlDR("start_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                    Format(sqlDR("end_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                    DateDiff(DateInterval.Minute, sqlDR("start_date"), sqlDR("end_date")) & vbTab &
                    sqlDR("lot_status")
                C1FlexGrid1.AddItem(insertString)
            Loop
            sqlDR.Close()

            DBClose()

            'show totals
            UpdateTotals()

            C1FlexGrid1.AutoSizeCols()
            C1FlexGrid1.Redraw = True

            thread_LoadingFormEnd
        End If

    End Sub

    Private Sub UpdateTotals()
        'clear existing totals
        C1FlexGrid1.Subtotal(AggregateEnum.Clear)

        'calculate subtotals (three levels, totals on every column)
        'Dim c As Integer
        'For c = 3 To C1FlexGrid1.Cols.Count - 1
        C1FlexGrid1.Subtotal(AggregateEnum.Sum, 0, -1, 6, "Grand Total")
        C1FlexGrid1.Subtotal(AggregateEnum.Average, 0, -1, 10, "Grand Total")
        C1FlexGrid1.Subtotal(AggregateEnum.Sum, 1, 2, 6, "Total for {0}")
        'Next

        'done, autosize columns to finish
        C1FlexGrid1.AutoSizeCols()

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub frm_SMT_Working_History_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus

        'timer_SMT_Working.Enabled = False

    End Sub

    Private Sub frm_SMT_Working_History_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        timer_SMT_Working.Enabled = True

    End Sub

    Private Sub frm_SMT_Working_History_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate

        timer_SMT_Working.Enabled = False

    End Sub

    Private Sub ToolStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub grid_Daily_List_MouseMove(sender As Object, e As MouseEventArgs) Handles C1FlexGrid1.MouseMove

        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then

            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub

    Private Sub grid_LotList_LostFocus(sender As Object, e As EventArgs) Handles C1FlexGrid1.LostFocus

        frm_Main.lb_Status.Text = String.Empty

    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged

        If NumericUpDown1.Value < 5 And Not NumericUpDown1.Value = 0 Then
            MsgBox("새로고침 시간은 5초이하로 설정 할 수 없습니다.",
                   MsgBoxStyle.Information,
                   form_Msgbox_String)
            NumericUpDown1.Value = 5
        End If

    End Sub

    Private Sub C1FlexGrid1_KeyDown(sender As Object, e As KeyEventArgs) Handles C1FlexGrid1.KeyDown

        If e.Shift Then
            md_ETC.cellCal(sender, Nothing)
        Else
            frm_Main.lb_Status.Text = String.Empty
        End If

    End Sub
End Class
'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-07-24
'##### 수정일자 : 2023-07-24
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : Baking 공정의 투입 및 완료시간을 관리한다.

'############################################################################################################
'############################################################################################################

Imports MySql.Data.MySqlClient

Public Class frm_Work_Baking

    Dim form_msgbox_string As String = "Baking 작업 투입 / 완료"

    Private Sub BAKING_WORK_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        Select Case e.KeyCode
            Case 115 'F4
                BTN_Start_1_Click(Nothing, Nothing)
            Case 116 'F5
                BTN_End_1_Click(Nothing, Nothing)
            Case 119 'F8
                BTN_Start_2_Click(Nothing, Nothing)
            Case 120 'F9
                BTN_End_2_Click(Nothing, Nothing)
        End Select

    End Sub

    Private Sub BAKING_WORK_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        BTN_End_1.Enabled = False
        BTN_End_2.Enabled = False
        Timer1.Enabled = True
        Timer1.Interval = 500
        Timer2.Enabled = False
        Timer2.Interval = 10000

        Working_List()

    End Sub

    Private Sub TB_Scan1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TB_Scan1.KeyDown

        If e.KeyCode = 13 Then

            Dim s_yj_no As String = String.Empty
            Dim s_lot_no As String = String.Empty

            If CheckBox1.Checked Then
                s_lot_no = TB_Scan1.Text
            Else
                '##### 스캔 정보 분리 #####
                Dim barcode_result() As String = Split(TB_Scan1.Text, "/")

                If Not UBound(barcode_result) > 1 Then
                    MsgBox("Scan 결과를 세부적으로 나눌 수 없습니다." & vbCrLf &
                           "다시 Scan하거나 Barcode를 확인하여 주십시오.", MsgBoxStyle.Information, form_msgbox_string)
                    TB_Scan1.SelectAll()
                    TB_Scan1.Focus()
                    Exit Sub
                End If

                s_yj_no = barcode_result(1)
                s_lot_no = barcode_result(0)
            End If

            '##### List에 등록된 Lot인지 확인 #####
            For i = 0 To ListBox1.Items.Count - 1
                If ListBox1.Items(i).ToString = TB_Scan1.Text Then
                    MsgBox("1호기에 이미 등록된 Lot입니다", MsgBoxStyle.Information, form_msgbox_string)
                    TB_Scan1.SelectAll()
                    TB_Scan1.Focus()
                    Exit Sub
                End If
            Next

            For i = 0 To ListBox2.Items.Count - 1
                If ListBox2.Items(i).ToString = TB_Scan1.Text Then
                    MsgBox("2호기에 이미 등록된 Lot입니다", MsgBoxStyle.Information, form_msgbox_string)
                    TB_Scan1.SelectAll()
                    TB_Scan1.Focus()
                    Exit Sub
                End If
            Next

            '##### Lot의 Baking 이력 및 정상 입고품 확인 #####
            'LOT_STATUS가 Incoming상태인지, Baking중인 상태인지
            If Lot_Baking_Check(s_lot_no) = True Then
                MsgBox("Baking 작업 등록된 Lot 입니다.", MsgBoxStyle.Information, form_msgbox_string)
                TB_Scan1.SelectAll()
                TB_Scan1.Focus()
                Exit Sub
            End If

            If Lot_Incoming_Check(s_lot_no) = False Then
                MsgBox("수입검사가 진행되지 않았거나" & vbCrLf &
                        "정상 입고된 Lot이 아닙니다.", MsgBoxStyle.Information, form_msgbox_string)
                TB_Scan1.SelectAll()
                TB_Scan1.Focus()
                Exit Sub
            End If

            '##### Listbox에 Lot을 등록 #####
            ListBox1.Items.Add(TB_Scan1.Text)
            Dim barcode_result2() As String = Split(TB_Scan1.Text, "/")
            Dim total_lot As Integer = 0
            For i = 0 To ListBox1.Items.Count - 1
                If Not barcode_result2(0) & "*" Like ListBox1.Items(i).ToString Then
                    total_lot += 1
                End If
            Next

            Label5.Text = "Total : " & total_lot & " Lot"
            TB_Scan1.Text = String.Empty

        End If

    End Sub

    Private Sub TB_Scan2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TB_Scan2.KeyDown

        If e.KeyCode = 13 Then

            Dim s_yj_no As String = String.Empty
            Dim s_lot_no As String = String.Empty

            If CheckBox2.Checked Then
                s_lot_no = TB_Scan2.Text
            Else
                '##### 스캔 정보 분리 #####
                Dim barcode_result() As String = Split(TB_Scan2.Text, "/")

                If Not UBound(barcode_result) > 1 Then
                    MsgBox("Scan 결과를 세부적으로 나눌 수 없습니다." & vbCrLf &
                           "다시 Scan하거나 Barcode를 확인하여 주십시오.", MsgBoxStyle.Information, form_msgbox_string)
                    TB_Scan2.SelectAll()
                    TB_Scan2.Focus()
                    Exit Sub
                End If

                s_yj_no = barcode_result(1)
                s_lot_no = barcode_result(0)
            End If

            '##### List에 등록된 Lot인지 확인 #####
            For i = 0 To ListBox2.Items.Count - 1
                If ListBox2.Items(i).ToString = TB_Scan2.Text Then
                    MsgBox("2호기에 이미 등록된 Lot입니다", MsgBoxStyle.Information, form_msgbox_string)
                    TB_Scan2.SelectAll()
                    TB_Scan2.Focus()
                    Exit Sub
                End If
            Next

            For i = 0 To ListBox1.Items.Count - 1
                If ListBox1.Items(i).ToString = TB_Scan2.Text Then
                    MsgBox("1호기에 이미 등록된 Lot입니다", MsgBoxStyle.Information, form_msgbox_string)
                    TB_Scan2.SelectAll()
                    TB_Scan2.Focus()
                    Exit Sub
                End If
            Next

            '##### Lot의 Baking 이력 및 정상 입고품 확인 #####
            'LOT_STATUS가 Incoming상태인지, Baking중인 상태인지
            If Lot_Baking_Check(s_lot_no) = True Then
                MsgBox("Baking 작업 등록된 Lot 입니다.", MsgBoxStyle.Information, form_msgbox_string)
                TB_Scan2.SelectAll()
                TB_Scan2.Focus()
                Exit Sub
            End If

            If Lot_Incoming_Check(s_lot_no) = False Then
                MsgBox("수입검사가 진행되지 않았거나" & vbCrLf &
                        "정상 입고된 Lot이 아닙니다.", MsgBoxStyle.Information, form_msgbox_string)
                TB_Scan2.SelectAll()
                TB_Scan2.Focus()
                Exit Sub
            End If

            '##### Listbox에 Lot을 등록 #####
            ListBox2.Items.Add(TB_Scan2.Text)
            Dim barcode_result2() As String = Split(TB_Scan2.Text, "/")
            Dim total_lot As Integer = 0
            For i = 0 To ListBox2.Items.Count - 1
                If Not barcode_result2(0) & "*" Like ListBox2.Items(i).ToString Then
                    total_lot += 1
                End If
            Next

            Label6.Text = "Total : " & total_lot & " Lot"
            TB_Scan2.Text = String.Empty

        End If

    End Sub

    Private Function Lot_Baking_Check(ByVal s_lot_no As String) As Boolean

        Lot_Baking_Check = False

        DBConnect()

        Dim strSQL As String = "select lot_no from work_baking where lot_no = '" & s_lot_no & "'"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Lot_Baking_Check = True
        Loop
        sqlDR.Close()

        DBClose()

        Return Lot_Baking_Check

    End Function

    Private Function Lot_Incoming_Check(ByVal s_lot_no As String) As Boolean

        Lot_Incoming_Check = False

        DBConnect()

        Dim strSQL As String = "select lot_no from incoming_inspect_result where lot_no = '" & s_lot_no & "'"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Lot_Incoming_Check = True
        Loop
        sqlDR.Close()

        DBClose()

        Return Lot_Incoming_Check

    End Function

    Private Function before_lot_check(ByVal checkListBox As ListBox) As Boolean

        before_lot_check = False

        DBConnect()

        For i = 0 To checkListBox.Items.Count - 1
            Dim printingLabel_Qty As Integer = 0
            Dim s_lot_no As String = checkListBox.Items(i).ToString.Split("/")(0)
            Dim strSQL As String = "select print_qty from history_label_print where lot_no = '" & s_lot_no & "'"
            strSQL += " and process_char = 'B' and print_method = '입고검사'"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                printingLabel_Qty = sqlDR("print_qty")
            Loop
            sqlDR.Close()

            Dim sameCount As Integer = 0
            For j = 0 To checkListBox.Items.Count - 1
                If checkListBox.Items(j).ToString Like s_lot_no & "*" Then
                    sameCount += 1
                End If
            Next

            If Not sameCount = printingLabel_Qty Then
                before_lot_check = True
                Exit For
            End If
        Next

        DBClose()

        Return before_lot_check

    End Function

    Private Sub BTN_Start_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Start_1.Click

        If tb_Worker1.Text = String.Empty Then
            MsgBox("작업자를 입력하여 주십시오.", MsgBoxStyle.Information, form_msgbox_string)
            tb_Worker1.SelectAll()
            tb_Worker1.Focus()
            Exit Sub
        End If

        If ListBox1.Items.Count = 0 Then
            MsgBox("등록된 Lot이 없습니다.", MsgBoxStyle.Information, form_msgbox_string)
            Exit Sub
        End If

        If before_lot_check(ListBox1) = True Then
            MsgBox("모두 투입되지 않은 Lot이 존재합니다." & vbCrLf & "Lot 수량을 확인하여 주십시오.",
                    MsgBoxStyle.Information, form_msgbox_string)
            Exit Sub
        End If

        If MsgBox("1호기 작업시작 등록을 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.No Then Exit Sub

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try

            TB_Start_1.Text = Format(Now, "yyyy-MM-dd HH:mm:ss")

            For i = 0 To ListBox1.Items.Count - 1

                Dim s_yj_no As String = String.Empty
                Dim s_lot_no As String = String.Empty

                Try
                    Dim barcode_result() As String = Split(ListBox1.Items(i).ToString, "/")

                    s_yj_no = barcode_result(1)
                    s_lot_no = barcode_result(0)
                Catch
                    s_yj_no = String.Empty
                    s_lot_no = ListBox1.Items(i).ToString
                End Try

                strSQL += "insert into work_baking(lot_no, work_status, yj_no, start_date, machine_no, label_info, baking_worker) values"
                strSQL += "('" & s_lot_no & "'"
                strSQL += ",'Start'"
                strSQL += ",'" & s_yj_no & "'"
                strSQL += ",'" & TB_Start_1.Text & "'"
                strSQL += ",1"
                strSQL += ",'" & ListBox1.Items(i).ToString & "'"
                strSQL += ",'" & tb_Worker1.Text & "');"

                strSQL += "update basic_lot_information set lot_status = 'Baking Start'"
                strSQL += " where lot_no = '" & s_lot_no & "';"
            Next

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_msgbox_string)
            Exit Sub
        End Try

        DBClose()

        MsgBox("작업시작 등록이 완료되었습니다.", MsgBoxStyle.Information, form_msgbox_string)
        Label16.Text = "Baking 1호기 작업중"
        BTN_End_1.Enabled = True
        BTN_Start_1.Enabled = False
        TB_Scan1.Enabled = False
        tb_Worker1.Enabled = False

    End Sub

    Private Sub Form_CLose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub Working_List()

        Label16.Text = "Baking 1호기 작업 대기중"
        Label1.Text = "Baking 2호기 작업 대기중"

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()

        DBConnect()

        Dim strSQL As String = "select start_date, label_info, baking_worker from work_baking where machine_no = 1"
        strSQL += " and work_status = 'Start'"
        strSQL += " order by yj_no"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            ListBox1.Items.Add(sqlDR("label_info"))
            TB_Start_1.Text = Format(sqlDR("start_date"), "yyyy-MM-dd HH:mm:ss")
            tb_Worker1.Text = sqlDR("baking_worker")
        Loop
        sqlDR.Close()

        strSQL = "select start_date, label_info, baking_worker from work_baking where machine_no = 2"
        strSQL += " and work_status = 'Start'"
        strSQL += " order by yj_no"

        sqlCmd = New MySqlCommand(strSQL, dbConnection1)
        sqlDR = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            ListBox2.Items.Add(sqlDR("label_info"))
            TB_Start_2.Text = Format(sqlDR("start_date"), "yyyy-MM-dd HH:mm:ss")
            tb_Worker2.Text = sqlDR("baking_worker")
        Loop
        sqlDR.Close()

        DBClose()

        If ListBox1.Items.Count > 0 Then
            BTN_Start_1.Enabled = False
            BTN_End_1.Enabled = True
            TB_Scan1.Enabled = False
            tb_Worker1.Enabled = False
            Label16.Text = "Baking 1호기 작업중"
            Dim total_lot_no As String = String.Empty
            For i = 0 To ListBox1.Items.Count - 1
                Dim now_lot As String = ListBox1.Items(i).ToString.Split("/")(0)
                If Not (total_lot_no Like "*" & now_lot & "*") Then
                    If total_lot_no = String.Empty Then
                        total_lot_no = now_lot
                    Else
                        total_lot_no += "," & now_lot
                    End If
                End If
            Next
            Label5.Text = "Total : " & UBound(total_lot_no.Split(",")) + 1 & " Lot"
        End If

        If ListBox2.Items.Count > 0 Then
            BTN_Start_2.Enabled = False
            BTN_End_2.Enabled = True
            TB_Scan2.Enabled = False
            tb_Worker2.Enabled = False
            Label1.Text = "Baking 2호기 작업중"
            Dim total_lot_no2 As String = String.Empty
            For i = 0 To ListBox2.Items.Count - 1
                Dim now_lot As String = ListBox2.Items(i).ToString.Split("/")(0)
                If Not (total_lot_no2 Like "*" & now_lot & "*") Then
                    If total_lot_no2 = String.Empty Then
                        total_lot_no2 = now_lot
                    Else
                        total_lot_no2 += "," & now_lot
                    End If
                End If
            Next
            Label6.Text = "Total : " & UBound(total_lot_no2.Split(",")) + 1 & " Lot"
        End If

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        If Label16.Text = "Baking 1호기 작업중" Then
            Label16.Text = String.Empty
        ElseIf Label16.Text = String.Empty Then
            Label16.Text = "Baking 1호기 작업중"
        End If

        If Label1.Text = "Baking 2호기 작업중" Then
            Label1.Text = String.Empty
        ElseIf Label1.Text = String.Empty Then
            Label1.Text = "Baking 2호기 작업중"
        End If

    End Sub

    Private Sub BTN_End_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_End_1.Click

        Dim min_hour As Date = DateAdd(DateInterval.Hour, 1, CDate(TB_Start_1.Text))
        Dim forceStop As String = "No"

        If Now < min_hour Then
            If MsgBox("작업시간이 충족되지 않았습니다." & vbCrLf & "강제 종료로 등록 하시겠습니까?",
                      MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                      form_msgbox_string) = MsgBoxResult.No Then Exit Sub
            forceStop = "Yes"
        End If

        If MsgBox("1호기 작업완료 등록을 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.No Then Exit Sub

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try

            TB_End_1.Text = Format(Now, "yyyy-MM-dd HH:mm:ss")

            For i = 0 To ListBox1.Items.Count - 1

                Dim s_yj_no As String = String.Empty
                Dim s_lot_no As String = String.Empty

                Try
                    Dim barcode_result() As String = Split(ListBox1.Items(i).ToString, "/")

                    s_yj_no = barcode_result(1)
                    s_lot_no = barcode_result(0)
                Catch
                    s_yj_no = String.Empty
                    s_lot_no = ListBox1.Items(i).ToString
                End Try

                strSQL += "update work_baking set end_date = '" & TB_End_1.Text & "'"
                strSQL += ", work_status = 'End'"
                strSQL += ", force_stop = '" & forceStop & "'"
                strSQL += " where lot_no = '" & s_lot_no & "';"

                strSQL += "update basic_lot_information set lot_status = 'Baking End'"
                strSQL += " where lot_no = '" & s_lot_no & "';"
            Next

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_msgbox_string)
            Exit Sub
        End Try

        DBClose()

        MsgBox("작업완료 등록이 완료되었습니다.", MsgBoxStyle.Information, form_msgbox_string)
        Label16.Text = "Baking 1호기 작업 대기중"
        BTN_End_1.Enabled = False
        BTN_Start_1.Enabled = True
        TB_Scan1.Enabled = True
        TB_Start_1.Text = String.Empty
        TB_End_1.Text = String.Empty
        tb_Worker1.Enabled = True

        ListBox1.Items.Clear()
        TB_Scan1.Focus()

    End Sub

    Private Sub BTN_Start_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Start_2.Click

        If tb_Worker2.Text = String.Empty Then
            MsgBox("작업자를 입력하여 주십시오.", MsgBoxStyle.Information, form_msgbox_string)
            tb_Worker2.SelectAll()
            tb_Worker2.Focus()
            Exit Sub
        End If

        If ListBox2.Items.Count = 0 Then
            MsgBox("등록된 Lot이 없습니다.", MsgBoxStyle.Information, form_msgbox_string)
            Exit Sub
        End If

        If before_lot_check(ListBox2) = True Then
            MsgBox("모두 투입되지 않은 Lot이 존재합니다." & vbCrLf & "Lot 수량을 확인하여 주십시오.",
                    MsgBoxStyle.Information, form_msgbox_string)
            Exit Sub
        End If

        If MsgBox("2호기 작업시작 등록을 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.No Then Exit Sub

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try

            TB_Start_2.Text = Format(Now, "yyyy-MM-dd HH:mm:ss")

            For i = 0 To ListBox2.Items.Count - 1

                Dim s_yj_no As String = String.Empty
                Dim s_lot_no As String = String.Empty

                Try
                    Dim barcode_result() As String = Split(ListBox2.Items(i).ToString, "/")

                    s_yj_no = barcode_result(1)
                    s_lot_no = barcode_result(0)
                Catch
                    s_yj_no = String.Empty
                    s_lot_no = ListBox2.Items(i).ToString
                End Try

                strSQL += "insert into work_baking(lot_no, work_status, yj_no, start_date, machine_no, label_info, baking_worker) values"
                strSQL += "('" & s_lot_no & "'"
                strSQL += ",'Start'"
                strSQL += ",'" & s_yj_no & "'"
                strSQL += ",'" & TB_Start_2.Text & "'"
                strSQL += ",2"
                strSQL += ",'" & ListBox2.Items(i).ToString & "'"
                strSQL += ",'" & tb_Worker2.Text & "');"

                strSQL += "update basic_lot_information set lot_status = 'Baking Start'"
                strSQL += " where lot_no = '" & s_lot_no & "';"
            Next

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_msgbox_string)
            Exit Sub
        End Try

        DBClose()

        MsgBox("작업시작 등록이 완료되었습니다.", MsgBoxStyle.Information, form_msgbox_string)
        Label1.Text = "Baking 2호기 작업중"
        BTN_End_2.Enabled = True
        BTN_Start_2.Enabled = False
        TB_Scan2.Enabled = False
        tb_Worker2.Enabled = False

    End Sub

    Private Sub BTN_End_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_End_2.Click

        Dim min_hour As Date = DateAdd(DateInterval.Hour, 1, CDate(TB_Start_2.Text))
        Dim forceStop As String = "No"

        If Now < min_hour Then
            If MsgBox("작업시간이 충족되지 않았습니다." & vbCrLf & "강제 종료로 등록 하시겠습니까?",
                      MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                      form_msgbox_string) = MsgBoxResult.No Then Exit Sub
            forceStop = "Yes"
        End If

        If MsgBox("2호기 작업완료 등록을 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.No Then Exit Sub

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try

            TB_End_2.Text = Format(Now, "yyyy-MM-dd HH:mm:ss")

            For i = 0 To ListBox2.Items.Count - 1

                Dim s_yj_no As String = String.Empty
                Dim s_lot_no As String = String.Empty

                Try
                    Dim barcode_result() As String = Split(ListBox2.Items(i).ToString, "/")

                    s_yj_no = barcode_result(1)
                    s_lot_no = barcode_result(0)
                Catch
                    s_yj_no = String.Empty
                    s_lot_no = ListBox2.Items(i).ToString
                End Try

                strSQL += "update work_baking set end_date = '" & TB_End_2.Text & "'"
                strSQL += ", work_status = 'End'"
                strSQL += ", force_stop = '" & forceStop & "'"
                strSQL += " where lot_no = '" & s_lot_no & "';"

                strSQL += "update basic_lot_information set lot_status = 'Baking End'"
                strSQL += " where lot_no = '" & s_lot_no & "';"
            Next

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_msgbox_string)
            Exit Sub
        End Try

        DBClose()

        MsgBox("작업완료 등록이 완료되었습니다.", MsgBoxStyle.Information, form_msgbox_string)
        Label1.Text = "Baking 2호기 작업 대기중"
        BTN_End_2.Enabled = False
        BTN_Start_2.Enabled = True
        TB_Scan2.Enabled = True
        TB_Start_2.Text = String.Empty
        TB_End_2.Text = String.Empty
        tb_Worker2.Enabled = True

        ListBox2.Items.Clear()
        TB_Scan2.Focus()

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        TB_Scan1.SelectAll()
        TB_Scan1.Focus()

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged

        TB_Scan2.SelectAll()
        TB_Scan2.Focus()

    End Sub

    Private Sub ListBox1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseDoubleClick

        If e.Button = MouseButtons.Left And
            Not (Label16.Text = "Baking 1호기 작업중" Or Label16.Text = String.Empty) Then
            Dim label_text As String = ListBox1.SelectedItem.ToString
            'MsgBox(label_text)
            If MsgBox(label_text & "를" & vbCrLf & "삭제 하시겠습니까?" _
                    , MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.No Then Exit Sub

            ListBox1.Items.Remove(ListBox1.SelectedItem)
        End If

    End Sub

    Private Sub ListBox2_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListBox2.MouseDoubleClick

        If e.Button = MouseButtons.Left And
            Not (Label1.Text = "Baking 2호기 작업중" Or Label1.Text = String.Empty) Then
            Dim label_text As String = ListBox2.SelectedItem.ToString
            'MsgBox(label_text)
            If MsgBox(label_text & "를" & vbCrLf & "삭제 하시겠습니까?" _
                    , MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.No Then Exit Sub

            ListBox2.Items.Remove(ListBox2.SelectedItem)
        End If

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        Working_List()

    End Sub

    Private Sub TB_Scan1_TextChanged(sender As Object, e As EventArgs) Handles TB_Scan1.TextChanged

    End Sub

    Private Sub btn_Reset_Data_Click(sender As Object, e As EventArgs) Handles btn_Reset_Data.Click

        Working_List()

    End Sub
End Class
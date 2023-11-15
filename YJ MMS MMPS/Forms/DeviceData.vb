Imports System.Xml
Imports C1.Win.C1FlexGrid
Imports CrystalDecisions.CrystalReports.Engine
Imports MySql.Data.MySqlClient

Public Class DeviceData

    Dim form_Msgbox_String As String = "DeviceData 등록"
    Dim writeReady As Boolean = False
    Dim newDeviceData As String

    Private Sub DeviceData_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        GridSetting()

    End Sub

    Private Sub GridSetting()

        With Grid_DeviceData
            .AllowEditing = True
            .AllowFiltering = True
            .AutoClipboard = True
            .AllowSorting = AllowSortingEnum.None
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 10
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            Grid_DeviceData(0, 0) = "No"
            Grid_DeviceData(0, 1) = "Code"
            Grid_DeviceData(0, 2) = "Machine No."
            Grid_DeviceData(0, 3) = "Feeder Slot No."
            Grid_DeviceData(0, 4) = "Maker"
            Grid_DeviceData(0, 5) = "Part No."
            Grid_DeviceData(0, 6) = "Specification"
            Grid_DeviceData(0, 7) = "Feeder SN"
            Grid_DeviceData(0, 8) = "대체품 수"
            Grid_DeviceData(0, 9) = "비고"
            .Cols(2).DataType = GetType(Integer)
            .Cols(2).Format = "#,##0"
            .Cols(3).DataType = GetType(Integer)
            .Cols(3).Format = "#,##0"
            .Cols(8).DataType = GetType(Integer)
            .Cols(8).Format = "#,##0"
            .Cols(8).ComboList = "..."
            .Cols(4).ComboList = Maker_List()
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(Grid_DeviceData.Cols.Count - 1).TextAlign = TextAlignEnum.LeftCenter
            .Cols(Grid_DeviceData.Cols.Count - 3).TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ExtendLastCol = True
        End With

    End Sub

    Private Sub Cb_customerName_DropDown(sender As Object, e As EventArgs) Handles Cb_customerName.DropDown

        Dim cb_old_string As String = Cb_customerName.Text

        Cb_customerName.Items.Clear()

        DBConnect()

        Dim strSQL As String = "select CUSTOMER_NAME from TB_CUSTOMER_LIST"
        strSQL += " order by CUSTOMER_NAME"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Cb_customerName.Items.Add(sqlDR("CUSTOMER_NAME"))
        Loop
        sqlDR.Close()

        DBClose()

        Cb_customerName.Text = cb_old_string

    End Sub

    Private Sub Cb_customerName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cb_customerName.SelectedIndexChanged

        Tb_customerCode.Text = String.Empty

        DBConnect()

        Dim strSQL As String = "select CUSTOMER_CODE from TB_CUSTOMER_LIST"
        strSQL += " where CUSTOMER_NAME = '" & Cb_customerName.Text & "'"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Tb_customerCode.Text = sqlDR("CUSTOMER_CODE")
        Loop
        sqlDR.Close()

        DBClose()

        If Btn_newCustomer.Text = "저장" Then
            Cb_customerName.DropDownStyle = ComboBoxStyle.DropDownList
            Btn_newCustomer.Text = "신규"
        End If

        Tb_modelCode.Text = String.Empty
        Cb_modelName.SelectedIndex = -1
        Cb_FactoryName.SelectedIndex = -1
        Cb_workLine.SelectedIndex = -1
        Cb_workSide.SelectedIndex = -1
        Grid_DeviceData.Rows.Count = 1

        writeReady = False

    End Sub

    Private Sub Cb_modelName_DropDown(sender As Object, e As EventArgs) Handles Cb_modelName.DropDown


        Dim cb_old_string As String = Cb_modelName.Text

        Cb_modelName.Items.Clear()

        DBConnect()

        Dim strSQL As String = "select MODEL_NAME from TB_MODEL_LIST"
        strSQL += " where CUSTOMER_CODE = '" & Tb_customerCode.Text & "'"
        strSQL += " order by MODEL_NAME"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Cb_modelName.Items.Add(sqlDR("MODEL_NAME"))
        Loop
        sqlDR.Close()

        DBClose()

        Cb_modelName.Text = cb_old_string

    End Sub

    Private Sub Btn_newCustomer_Click(sender As Object, e As EventArgs) Handles Btn_newCustomer.Click

        If Btn_newCustomer.Text = "신규" Then
            Cb_customerName.DropDownStyle = ComboBoxStyle.DropDown
            Cb_customerName.SelectedIndex = -1
            Btn_newCustomer.Text = "저장"
            Cb_modelName.DropDownStyle = ComboBoxStyle.DropDownList
            Cb_modelName.SelectedIndex = -1
            Tb_modelCode.Text = String.Empty
            Cb_workLine.SelectedIndex = -1
            Cb_workSide.SelectedIndex = -1
            Tb_customerCode.Text = Customer_Code_Make()
            Cb_customerName.Focus()
        Else
            Dim customerName As String = Trim(Cb_customerName.Text)
            If customerName = String.Empty Then Exit Sub
            If MsgBox("고객사명 " & customerName & "(을)를 신규로 등록 하시겟습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                  form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

            If NewCustomerAdd() = False Then Exit Sub
            Cb_customerName.DropDownStyle = ComboBoxStyle.DropDownList
            Btn_newCustomer.Text = "신규"
            Cb_customerName.Items.Add(customerName)
            Cb_customerName.SelectedIndex = Cb_customerName.Items.Count - 1
            MsgBox("신규 고객사 등록 완료.", MsgBoxStyle.Information, form_Msgbox_String)
        End If

    End Sub

    Private Function NewCustomerAdd() As Boolean

        Dim result As Boolean = True

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = DBConnect1.BeginTransaction

        Try

            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            strSQL = "insert into TB_CUSTOMER_LIST(CUSTOMER_CODE, CUSTOMER_NAME, WRITE_DATE, WRITE_ID) values"
            strSQL += "('" & Tb_customerCode.Text & "'"
            strSQL += ",'" & Cb_customerName.Text & "'"
            strSQL += ",'" & write_date & "'"
            strSQL += ",'" & Mainform.login_user & "');"

            strSQL += "insert into TB_SUB_CODE(WRITE_DATE, SUB_CODE, SUB_CODE_NAME, MAIN_CODE, SUB_CODE_NOTE) values"
            strSQL += "('" & write_date & "'"
            strSQL += ",'" & Tb_customerCode.Text & "'"
            strSQL += ",'" & Cb_customerName.Text & "'"
            strSQL += ",'MC0001'"
            strSQL += ",'');"

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, DBConnect1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            result = False
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
        End Try

        DBClose()

        Return result

    End Function

    Private Function Customer_Code_Make() As String

        Dim lastCode As String = String.Empty

        DBConnect()

        Dim strSQL As String = "select CUSTOMER_CODE from TB_CUSTOMER_LIST order by CUSTOMER_CODE desc limit 1"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            lastCode = sqlDR("CUSTOMER_CODE")
        Loop
        sqlDR.Close()

        DBClose()

        If lastCode = String.Empty Then
            lastCode = "CC0001"
        Else
            Dim addOrder As Integer = CInt(lastCode.Substring(2, 4)) + 1
            lastCode = "CC" & Format(addOrder, "0000")
        End If

        Return lastCode

    End Function

    Dim newModel As Boolean = False
    Dim newModelSave As Boolean = False

    Private Sub Btn_newModel_Click(sender As Object, e As EventArgs) Handles Btn_newModel.Click

        If Customer_Code_Check() = False Then
            MsgBox("등록되지 않은 고객사 입니다.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If Btn_newModel.Text = "신규" Then
            Cb_modelName.DropDownStyle = ComboBoxStyle.DropDown
            Cb_modelName.SelectedIndex = -1
            Btn_newModel.Text = "저장"
            Cb_workLine.SelectedIndex = -1
            Cb_workSide.SelectedIndex = -1
            Tb_modelCode.Text = Model_Code_Make()
            Cb_modelName.Focus()
            newModel = True
        Else
            Dim modelName As String = Trim(Cb_modelName.Text)
            If modelName = String.Empty Then Exit Sub
            If MsgBox("모델명 " & modelName & "(을)를 신규로 등록 하시겟습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                  form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

            '중복검사
            If ExsistModelCheck() = True Then
                MsgBox("중복된 모델명입니다.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, form_Msgbox_String)
                Exit Sub
            End If

            '저장하기
            If NewModelAdd() = False Then Exit Sub
            Cb_modelName.DropDownStyle = ComboBoxStyle.DropDownList
            Btn_newModel.Text = "신규"
            Cb_modelName.Items.Add(modelName)
            Cb_modelName.SelectedIndex = Cb_modelName.Items.Count - 1
            MsgBox("신규 모델 등록 완료.", MsgBoxStyle.Information, form_Msgbox_String)
            newModelSave = True
        End If

    End Sub

    Private Function ExsistModelCheck() As Boolean

        DBConnect()

        ExsistModelCheck = False

        Dim strSQL As String = "select * from TB_MODEL_LIST where MODEL_NAME = '" & Cb_modelName.Text & "' and CUSTOMER_CODE = '" & Tb_customerCode.Text & "';"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            ExsistModelCheck = True
        Loop
        sqlDR.Close()

        DBClose()

        Return ExsistModelCheck

    End Function

    Private Function Customer_Code_Check() As Boolean

        Dim lastCode As Boolean = False

        DBConnect()

        Dim strSQL As String = "select CUSTOMER_CODE from TB_CUSTOMER_LIST"
        strSQL += " where CUSTOMER_CODE = '" & Tb_customerCode.Text & "'"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            lastCode = True
        Loop
        sqlDR.Close()

        DBClose()

        Return lastCode

    End Function

    Private Function Model_Code_Make() As String

        Dim lastCode As String = String.Empty

        DBConnect()

        Dim strSQL As String = "select MODEL_CODE from TB_MODEL_LIST"
        strSQL += " order by MODEL_CODE desc limit 1"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            lastCode = sqlDR("MODEL_CODE")
        Loop
        sqlDR.Close()

        DBClose()

        If lastCode = String.Empty Then
            lastCode = "MC0001"
        Else
            Dim addOrder As Integer = CInt(lastCode.Substring(2, 4)) + 1
            lastCode = "MC" & Format(addOrder, "0000")
        End If

        Return lastCode

    End Function

    Private Function NewModelAdd() As Boolean

        Dim result As Boolean = True

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = DBConnect1.BeginTransaction

        Try

            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            strSQL = "insert into TB_MODEL_LIST(MODEL_CODE, MODEL_NAME, CUSTOMER_CODE, WRITE_dATE, WRITE_ID) values"
            strSQL += "('" & Tb_modelCode.Text & "'"
            strSQL += ",'" & Cb_modelName.Text & "'"
            strSQL += ",'" & Tb_customerCode.Text & "'"
            strSQL += ",'" & write_date & "'"
            strSQL += ",'" & Mainform.login_user & "');"

            strSQL += "insert into TB_LAST_CODE(WRITE_DATE, LAST_CODE, LAST_CODE_NAME, SUB_CODE, MAIN_CODE, LAST_CODE_NOTE) values"
            strSQL += "('" & write_date & "'"
            strSQL += ",'" & Tb_modelCode.Text & "'"
            strSQL += ",'" & Cb_modelName.Text & "'"
            strSQL += ",'" & Tb_customerCode.Text & "'"
            strSQL += ",'MC0001'"
            strSQL += ",'');"

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, DBConnect1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            result = False
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
        End Try

        DBClose()

        Return result

    End Function

    Private Sub Cb_modelName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cb_modelName.SelectedIndexChanged

        Tb_modelCode.Text = String.Empty

        DBConnect()

        Dim strSQL As String = "select MODEL_CODE from TB_MODEL_LIST"
        strSQL += " where MODEL_NAME = '" & Cb_modelName.Text & "'"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Tb_modelCode.Text = sqlDR("MODEL_CODE")
        Loop
        sqlDR.Close()

        DBClose()

        If Btn_newModel.Text = "저장" Then
            Cb_modelName.DropDownStyle = ComboBoxStyle.DropDownList
            Btn_newModel.Text = "신규"
        End If

        Cb_FactoryName.SelectedIndex = -1
        Cb_workLine.SelectedIndex = -1
        Cb_workSide.SelectedIndex = -1
        Grid_DeviceData.Rows.Count = 1

        writeReady = False

    End Sub

    Public Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        If Tb_customerCode.Text = String.Empty Then
            MsgBox("고객사를 선택하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If Tb_modelCode.Text = String.Empty Then
            MsgBox("Model을 선택하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If Cb_FactoryName.Text = String.Empty Then
            MsgBox("공장을 선택하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If Cb_workLine.Text = String.Empty Then
            MsgBox("작업라인을 선택하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If Cb_workSide.Text = String.Empty Then
            MsgBox("작업면을 선택하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        newModel = False
        newModelSave = False

        Grid_DeviceData.Redraw = False
        Grid_DeviceData.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call USP_DEVICE_DATA(0"
        strSQL += ",'" & Tb_modelCode.Text & "'"
        strSQL += ",'" & Cb_workLine.Text & "'"
        strSQL += ",'" & Cb_workSide.Text & "'"
        strSQL += ",'" & Cb_FactoryName.Text & "'"
        strSQL += ",null"
        strSQL += ",'" & Cb_customerName.Text & "')"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert As String = String.Empty
            insert = Grid_DeviceData.Rows.Count &
                vbTab & sqlDR("DD_CODE") &
                vbTab & sqlDR("MACHINE_NO") &
                vbTab & sqlDR("FEEDER_NO") &
                vbTab & sqlDR("PART_MAKER") &
                vbTab & sqlDR("PART_NO") &
                vbTab & sqlDR("PART_SPECIFICATION") &
                vbTab & sqlDR("FEEDER_SN") &
                vbTab & sqlDR("MATCH_COUNT") &
                vbTab & sqlDR("DD_NOTE")
            Grid_DeviceData.AddItem(insert)
            newDeviceData = sqlDR("DD_MAIN_NO")
        Loop
        sqlDR.Close()

        DBClose()

        Grid_DeviceData.AutoSizeCols()
        Grid_DeviceData.Redraw = True

        If Grid_DeviceData.Rows.Count = 1 Then
            MsgBox("조회된 내용이 없습니다." & vbCrLf & "신규 데이터를 작성하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            newDeviceData = "DD-" & Format(Now, "yyyyMMddHHmmss")
        End If

        writeReady = True

    End Sub

    Private Sub Grid_DeviceData_RowColChange(sender As Object, e As EventArgs) Handles Grid_DeviceData.RowColChange

        If Grid_DeviceData.Row > 0 Then
            Select Case Grid_DeviceData.Col
                Case 1
                    Grid_DeviceData.AllowEditing = False
                Case Else
                    If Not IsNothing(Grid_DeviceData(Grid_DeviceData.Row, 0)) Then
                        If Grid_DeviceData(Grid_DeviceData.Row, 0).ToString = "D" Then
                            Grid_DeviceData.AllowEditing = False
                        Else
                            Grid_DeviceData.AllowEditing = True
                        End If
                    End If
            End Select
        End If

    End Sub

    Private Sub Grid_DeviceData_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_DeviceData.MouseClick

        Dim row As Integer = Grid_DeviceData.MouseRow

        If e.Button = MouseButtons.Right And row > -1 Then
            If writeReady = True And
                    Not Tb_customerCode.Text = String.Empty And
                    Not Tb_modelCode.Text = String.Empty And
                    Not Cb_workLine.Text = String.Empty And
                    Not Cb_workSide.Text = String.Empty And
                    Not Cb_FactoryName.Text = String.Empty Then
                BTN_deleteLine.Enabled = True
                BTN_Save.Enabled = True
                If row = 0 Then
                    BTN_deleteLine.Enabled = False
                    BTN_Save.Enabled = False
                Else
                    If Grid_DeviceData(row, 0).ToString = "D" Then
                        BTN_deleteLine.Text = "삭제취소"
                    Else
                        BTN_deleteLine.Text = "삭제"
                    End If
                End If
                Grid_DeviceData.Row = row
                CMS_gridMenu.Show(Grid_DeviceData, New Point(e.X, e.Y))
            End If
        End If

    End Sub

    Private Sub Cb_workLine_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cb_workLine.SelectedIndexChanged

        writeReady = False
        Grid_DeviceData.Rows.Count = 1

    End Sub

    Private Sub Cb_workSide_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cb_workSide.SelectedIndexChanged

        writeReady = False
        Grid_DeviceData.Rows.Count = 1

    End Sub

    Private Sub Grid_Index_Sorting()

        For i = 1 To Grid_DeviceData.Rows.Count - 1
            If IsNumeric(Grid_DeviceData(i, 0)) Then
                Grid_DeviceData(i, 0) = i
            End If
        Next

    End Sub

    Private Sub BTN_newLine_Click(sender As Object, e As EventArgs) Handles BTN_newLine.Click

        Dim insertRow As Integer = 0
        If Grid_DeviceData.Row < 1 Then
            insertRow = 1
        Else
            insertRow = Grid_DeviceData.Row + 1
        End If

        Grid_DeviceData.Redraw = False

        Grid_DeviceData.AddItem("N" & vbTab & DeviceDataCode(), insertRow)
        Grid_DeviceData.Rows(insertRow).StyleNew.ForeColor = Color.Blue

        Grid_DeviceData.AutoSizeCols()

        Grid_DeviceData.Redraw = True

    End Sub

    Private Function DeviceDataCode() As String

        Dim newCode As String = String.Empty

        DBConnect()

        Dim strSQL As String = "call USP_DEVICE_DATA(1"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null)"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            newCode = sqlDR("DD_CODE")
        Loop
        sqlDR.Close()

        DBClose()

        If newCode = String.Empty Then
            newCode = "DD00000001"
        Else
            Dim sortNo As Integer = newCode.Substring(2, 8)
            newCode = "DD" & Format(sortNo + 1, "00000000")
        End If

        Dim last_code() As Integer = {}
        ReDim last_code(Grid_DeviceData.Rows.Count - 1)

        For i = 1 To Grid_DeviceData.Rows.Count - 1
            last_code(i) = Grid_DeviceData(i, 1).ToString.Substring(2, 8)
        Next

        If Not last_code.Length = 0 Then
            Dim first_code As String = newCode.Substring(0, 2)
            Dim second_code As Integer = newCode.Substring(2, 8)
            If second_code <= last_code.Max Then
                newCode = first_code + Format(last_code.Max + 1, "00000000")
            End If
        End If

        Return newCode

    End Function

    Private Sub BTN_deleteLine_Click(sender As Object, e As EventArgs) Handles BTN_deleteLine.Click

        Dim row As Integer = Grid_DeviceData.Row

        If Grid_DeviceData(row, 0).ToString = "D" Then
            Grid_DeviceData.Rows(row).StyleNew.ForeColor = Color.Black
            Grid_DeviceData(row, 0) = row
        ElseIf Grid_DeviceData(row, 0).ToString = "N" Then
            Grid_DeviceData.Rows.Remove(row)
        Else
            Grid_DeviceData(row, 0) = "D"
            Grid_DeviceData.Rows(row).StyleNew.ForeColor = Color.DarkGray
        End If

    End Sub

    Private Function Maker_List() As String

        Dim makerList As String = String.Empty

        DBConnect()

        Dim strSQL As String = "call USP_MAKER_LIST"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            If makerList = String.Empty Then
                makerList = sqlDR("SUB_CODE_NAME")
            Else
                makerList += "|" & sqlDR("SUB_CODE_NAME")
            End If
        Loop
        sqlDR.Close()

        DBClose()

        Return makerList

    End Function

    Private Sub Grid_DeviceData_ComboDropDown(sender As Object, e As RowColEventArgs) Handles Grid_DeviceData.ComboDropDown

        If e.Col = 4 Then
            Grid_DeviceData.Cols(4).ComboList = Maker_List()
        End If

    End Sub

    Private Sub BTN_Save_Click(sender As Object, e As EventArgs) Handles BTN_Save.Click

        If newModel = True Then
            If newModelSave = False Then
                MsgBox("신규 모델을 먼저 저장하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
                Exit Sub
            End If
        End If

        If MsgBox("저장 하시겠습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                  form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = DBConnect1.BeginTransaction

        Try

            Dim writeDate As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            For i = 1 To Grid_DeviceData.Rows.Count - 1
                If Grid_DeviceData(i, 0).ToString = "N" Then
                    strSQL += "insert into TB_DEVICE_DATA(DD_CODE, MODEL_CODE, FACTORY_NAME, LINE_NAME, WORK_SIDE"
                    strSQL += ", MACHINE_NO, FEEDER_NO, PART_MAKER, PART_NO, PART_SPECIFICATION, DD_NOTE, WRITE_ID, WRITE_DATE, DD_MAIN_NO, FEEDER_SN, FEEDER_DATE) values"
                    strSQL += "('" & Grid_DeviceData(i, 1) & "'"
                    strSQL += ",'" & Tb_modelCode.Text & "'"
                    strSQL += ",'" & Cb_FactoryName.Text & "'"
                    strSQL += ",'" & Cb_workLine.Text & "'"
                    strSQL += ",'" & Cb_workSide.Text & "'"
                    strSQL += ",'" & Grid_DeviceData(i, 2) & "'"
                    strSQL += ",'" & Grid_DeviceData(i, 3) & "'"
                    strSQL += ",'" & Grid_DeviceData(i, 4) & "'"
                    strSQL += ",'" & Grid_DeviceData(i, 5) & "'"
                    strSQL += ",'" & Grid_DeviceData(i, 6) & "'"
                    strSQL += ",'" & Grid_DeviceData(i, 9) & "'"
                    strSQL += ",'" & Mainform.login_user & "'"
                    strSQL += ",'" & writeDate & "'"
                    strSQL += ",'" & newDeviceData & "'"
                    strSQL += ",'" & Grid_DeviceData(i, 7) & "'"
                    strSQL += ",'" & writeDate & "');"
                ElseIf Grid_DeviceData(i, 0).ToString = "M" Then
                    strSQL += "update TB_DEVICE_DATA set MACHINE_NO = '" & Grid_DeviceData(i, 2) & "'"
                    strSQL += ", FEEDER_NO = '" & Grid_DeviceData(i, 3) & "'"
                    strSQL += ", PART_MAKER = '" & Grid_DeviceData(i, 4) & "'"
                    strSQL += ", PART_NO = '" & Grid_DeviceData(i, 5) & "'"
                    strSQL += ", PART_SPECIFICATION = '" & Grid_DeviceData(i, 6) & "'"
                    strSQL += ", DD_NOTE = '" & Grid_DeviceData(i, 9) & "'"
                    strSQL += ", WRITE_ID = '" & Mainform.login_user & "'"
                    strSQL += ", WRITE_DATE = '" & writeDate & "'"
                    strSQL += ", FEEDER_SN = '" & Grid_DeviceData(i, 7) & "'"
                    strSQL += ", FEEDER_DATE = '" & writeDate & "'"
                    strSQL += " where DD_CODE = '" & Grid_DeviceData(i, 1) & "';"
                    strSQL += "update TB_DEVICE_DATA set PART_SPECIFICATIOn = '" & Grid_DeviceData(i, 6) & "'"
                    strSQL += " where PART_MAKER = '" & Grid_DeviceData(i, 4) & "'"
                    strSQL += " and PART_NO = '" & Grid_DeviceData(i, 5) & "';"
                ElseIf Grid_DeviceData(i, 0).ToString = "D" Then
                    strSQL += "delete from TB_DEVICE_DATA"
                    strSQL += " where DD_CODE = '" & Grid_DeviceData(i, 1) & "';"
                End If
            Next

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, DBConnect1)
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

    Private Sub Grid_DeviceData_AfterEdit(sender As Object, e As RowColEventArgs) Handles Grid_DeviceData.AfterEdit

        If beforeText = Grid_DeviceData(e.Row, e.Col) Then Exit Sub

        Grid_DeviceData.Redraw = False

        If IsNumeric(Grid_DeviceData(e.Row, 0)) Then
            Grid_DeviceData(e.Row, 0) = "M"
            Grid_DeviceData.Rows(e.Row).StyleNew.ForeColor = Color.Red
        End If

        Select Case e.Col
            Case 2, 3
                If Not IsNothing(Grid_DeviceData(e.Row, 2)) And
                    Not IsNothing(Grid_DeviceData(e.Row, 3)) Then
                    Dim exist As Boolean = False
                    For i = 1 To Grid_DeviceData.Rows.Count - 1
                        If Not i = e.Row Then
                            If Grid_DeviceData(e.Row, 2) = Grid_DeviceData(i, 2) And
                                Grid_DeviceData(e.Row, 3) = Grid_DeviceData(i, 3) Then
                                exist = True
                                Exit For
                            End If
                        End If
                    Next

                    If exist = True Then
                        e.Cancel = True
                        MsgBox("Machine No와 Feeder No가 중복되었습니다.", MsgBoxStyle.Information, form_Msgbox_String)
                    End If
                End If
            Case 5
                'Grid_DeviceData(e.Row, 5) = Barcode_Split(Grid_DeviceData(e.Row, 5), Grid_DeviceData(e.Row, 4)).ToString.Split("@")(0)
                'Grid_DeviceData(e.Row, 6) = Part_Spec_Load(Grid_DeviceData(e.Row, 5))
                Dim partInfo As String = loadMakerSpec(Grid_DeviceData(e.Row, 5))
                Grid_DeviceData(e.Row, 4) = partInfo.Split("@")(0)
                Grid_DeviceData(e.Row, 6) = partInfo.Split("@")(1)
        End Select

        Grid_DeviceData.AutoSizeCols()

        Grid_DeviceData.Redraw = True

    End Sub

    Private Function Part_Spec_Load(ByVal partNo As String) As String

        Dim partSpec As String = String.Empty

        DBConnect()

        Try
            Dim strSQL As String = "call USP_DEVICE_DATA(2"
            strSQL += ",null"
            strSQL += ",null"
            strSQL += ",null"
            strSQL += ",null"
            strSQL += ",'" & partNo & "'"
            strSQL += ",null)"

            Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                partSpec = sqlDR("PART_SPECIFICATION")
            Loop
            sqlDR.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, form_Msgbox_String)
        End Try

        DBClose()

        Return partSpec

    End Function

    Dim beforeText As String

    Private Sub Grid_DeviceData_BeforeEdit(sender As Object, e As RowColEventArgs) Handles Grid_DeviceData.BeforeEdit

        beforeText = Grid_DeviceData(e.Row, e.Col)

    End Sub

    Private Sub BTN_SM_Feeder_Click(sender As Object, e As EventArgs) Handles BTN_SM_Feeder.Click

        If writeReady = False Then
            MsgBox("조회된 내용이 없습니다." & vbCrLf & "목록 선택 후 조회를 먼저 실행 하십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        'xml파일을 불러와서 보여준다.

        Dim openFileDialog1 As New OpenFileDialog

        openFileDialog1.FileName = String.Empty

        openFileDialog1.Filter = "XML Files (*.xml)|*.xml"
        openFileDialog1.ShowDialog()

        If openFileDialog1.FileName = String.Empty Then Exit Sub

        Grid_DeviceData.Redraw = False

        Dim doc = New XmlDocument
        doc.Load(openFileDialog1.FileName)

        Dim root As XmlNode = doc.DocumentElement

        Dim devices As XmlNodeList = root.SelectNodes("PCBDATA/MACHINE/DEVICES/FEEDERBASE")
        Dim machine_no As Integer = 1
        If IsNumeric(Grid_DeviceData(Grid_DeviceData.Rows.Count - 1, 2)) Then
            machine_no = CInt(Grid_DeviceData(Grid_DeviceData.Rows.Count - 1, 2)) + 1
        Else
            If Grid_DeviceData(Grid_DeviceData.Rows.Count - 1, 2) = String.Empty Then Grid_DeviceData.Rows.Count = 1
        End If
        Dim now_base_no As Integer = 1

        For Each base_no As XmlNode In devices
            Console.WriteLine(base_no.ChildNodes.ToString)
            Dim exist_StickFeeder As XmlNodeList = base_no.SelectNodes("STICKFEEDER")
            Dim exist_TapeFeeder As XmlNodeList = base_no.SelectNodes("TAPEFEEDER")
            Dim exist_TrayFeeder As XmlNodeList = base_no.SelectNodes("TRAYFEEDER")

            If Not now_base_no = base_no.Attributes("ID").Value Then 'Feeder Base No.
                machine_no += 1
            End If

            If exist_TapeFeeder.Count > 0 Then
                Dim selectFeeder As XmlNodeList = base_no.SelectNodes("TAPEFEEDER")
                For Each feeder_info As XmlNode In selectFeeder
                    Dim partInfo As String = loadMakerSpec(feeder_info.SelectSingleNode("TAPE").Attributes("PARTNAME").Value)
                    Dim insert_string As String = "N" & vbTab &
                          DeviceDataCode() & vbTab &
                          machine_no & vbTab &
                          feeder_info.SelectSingleNode("INSTALL").Attributes("SLOT").Value & vbTab &
                          partInfo.Split("@")(0) & vbTab &
                          feeder_info.SelectSingleNode("TAPE").Attributes("PARTNAME").Value & vbTab &
                          partInfo.Split("@")(1) & vbTab &
                          String.Empty & vbTab &
                          String.Empty & vbTab &
                          String.Empty
                    Grid_DeviceData.AddItem(insert_string)
                    Grid_DeviceData.Rows(Grid_DeviceData.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
                Next
            End If

            If exist_StickFeeder.Count > 0 Then
                Dim selectFeeder As XmlNodeList = base_no.SelectNodes("STICKFEEDER")
                For Each feeder_info As XmlNode In selectFeeder
                    Dim partInfo As String = loadMakerSpec(feeder_info.SelectSingleNode("STICK").Attributes("PARTNAME").Value)
                    Dim insert_string As String = "N" & vbTab &
                          DeviceDataCode() & vbTab &
                          machine_no & vbTab &
                          feeder_info.SelectSingleNode("INSTALL").Attributes("SLOT").Value & vbTab &
                          partInfo.Split("@")(0) & vbTab &
                          feeder_info.SelectSingleNode("STICK").Attributes("PARTNAME").Value & vbTab &
                          partInfo.Split("@")(1) & vbTab &
                          String.Empty & vbTab &
                          String.Empty & vbTab &
                          String.Empty
                    Grid_DeviceData.AddItem(insert_string)
                    Grid_DeviceData.Rows(Grid_DeviceData.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
                Next
            End If

            '트레이 부분은 좀더 구체적으로 들어가야한다.
            'If exist_TrayFeeder.Count > 0 Then
            '    Dim selectFeeder As XmlNodeList = base_no.SelectNodes("TRAYFEEDER")
            '    For Each feeder_info As XmlNode In selectFeeder
            '        Dim partInfo As String = loadMakerSpec(feeder_info.SelectSingleNode("TRAY").Attributes("PARTNAME").Value)
            '        Dim insert_string As String = "N" & vbTab &
            '              DeviceDataCode() & vbTab &
            '              machine_no & vbTab &
            '              feeder_info.SelectSingleNode("INSTALL").Attributes("SLOT").Value & vbTab &
            '              partInfo.Split("@")(0) & vbTab &
            '              feeder_info.SelectSingleNode("TRAY").Attributes("PARTNAME").Value & vbTab &
            '              partInfo.Split("@")(1) & vbTab &
            '              String.Empty & vbTab &
            '              String.Empty & vbTab &
            '              String.Empty
            '        Grid_DeviceData.AddItem(insert_string)
            '        Grid_DeviceData.Rows(Grid_DeviceData.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
            '    Next
            'End If
        Next

        Grid_DeviceData.AutoSizeCols()
        Grid_DeviceData.AutoSizeRows(1, 0, Grid_DeviceData.Rows.Count - 1, Grid_DeviceData.Cols.Count - 1, 0, AutoSizeFlags.None)
        Grid_DeviceData.Redraw = True

        MsgBox("SAMSUNG(HANWHA) Feeder List(*.xml)를(을) 불러 왔습니다.", MsgBoxStyle.Information, form_Msgbox_String)

    End Sub

    Private Sub BTN_DataPrint_Click(sender As Object, e As EventArgs) Handles BTN_DataPrint.Click

        Dim strSQL As String = String.Empty

        'MySQL DB에서 정보를 불러오기전에 기존 내용 삭제
        MDBConnect()

        Dim sqlTran_MDB As OleDb.OleDbTransaction
        Dim sqlCmd_MDB As OleDb.OleDbCommand

        sqlTran_MDB = MDBConnect1.BeginTransaction

        Try
            '기존 저장 데이터를 삭제
            strSQL = "delete from DEVICE_DATA_PRT;"

            sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, MDBConnect1)
            sqlCmd_MDB.Transaction = sqlTran_MDB
            sqlCmd_MDB.ExecuteNonQuery()

            strSQL = "delete from DEVICE_DATA_SUB;"

            sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, MDBConnect1)
            sqlCmd_MDB.Transaction = sqlTran_MDB
            sqlCmd_MDB.ExecuteNonQuery()

            sqlTran_MDB.Commit()

        Catch ex As OleDb.OleDbException
            sqlTran_MDB.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        '새로운 데이터를 MySQL에서 가져온다.
        sqlTran_MDB = MDBConnect1.BeginTransaction

        Try
            DBConnect()

            strSQL = "call USP_DEVICE_DATA(3"
            strSQL += ",'" & Tb_modelCode.Text & "'"
            strSQL += ",'" & Cb_workLine.Text & "'"
            strSQL += ",'" & Cb_workSide.Text & "'"
            strSQL += ",'" & Cb_FactoryName.Text & "'"
            strSQL += ",null"
            strSQL += ",null)"

            Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader
            '''' 서브도 집어 넣어야 된다!!!!
            Do While sqlDR.Read
                strSQL = "insert into DEVICE_DATA_PRT(DD_CODE, CUSTOMER_CODE, CUSTOMER_NAME"
                strSQL += ", MODEL_CODE, MODEL_NAME, LINE_NAME, SIDE, MACHINE_NO, FEEDER_NO, PART_MAKER, PART_NO"
                strSQL += ", SPECIFICATION, DD_NOTE, FEEDER_SN, MATCH_COUNT, DD_MAIN_NO) values"
                strSQL += "('" & sqlDR("DD_CODE") & "'"
                strSQL += ",'" & Tb_customerCode.Text & "'"
                strSQL += ",'" & Cb_customerName.Text & "'"
                strSQL += ",'" & Tb_modelCode.Text & "'"
                strSQL += ",'" & Cb_modelName.Text & "'"
                strSQL += ",'" & Cb_workLine.Text & "'"
                strSQL += ",'" & Cb_workSide.Text & "'"
                strSQL += ",'" & sqlDR("MACHINE_NO") & "'"
                strSQL += ",'" & sqlDR("FEEDER_NO") & "'"
                strSQL += ",'" & sqlDR("PART_MAKER") & "'"
                strSQL += ",'" & sqlDR("PART_NO") & "'"
                strSQL += ",'" & sqlDR("PART_SPECIFICATION") & "'"
                strSQL += ",'" & sqlDR("DD_NOTE") & "'"
                strSQL += ",'" & String.Empty & "'"
                strSQL += ",'" & sqlDR("MATCH_COUNT") & "'"
                strSQL += ",'" & sqlDR("DD_MAIN_NO") & "');"

                sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, MDBConnect1)
                sqlCmd_MDB.Transaction = sqlTran_MDB
                sqlCmd_MDB.ExecuteNonQuery()
            Loop
            sqlDR.Close()

            strSQL = "call USP_DEVICE_DATA(4"
            strSQL += ",'" & Tb_modelCode.Text & "'"
            strSQL += ",'" & Cb_workLine.Text & "'"
            strSQL += ",'" & Cb_workSide.Text & "'"
            strSQL += ",'" & Cb_FactoryName.Text & "'"
            strSQL += ",null"
            strSQL += ",null)"

            sqlCmd = New MySqlCommand(strSQL, DBConnect1)
            sqlDR = sqlCmd.ExecuteReader
            '''' 서브도 집어 넣어야 된다!!!!
            Do While sqlDR.Read
                strSQL = "insert into DEVICE_DATA_SUB(DD_CODE, DS_CODE, PART_MAKER, PART_NO, DD_NOTE) values"
                strSQL += "('" & sqlDR("DD_CODE") & "'"
                strSQL += ",'" & sqlDR("DS_CODE") & "'"
                strSQL += ",'" & sqlDR("PART_MAKER") & "'"
                strSQL += ",'" & sqlDR("PART_NO") & "'"
                strSQL += ",'" & sqlDR("DS_NOTE") & "');"

                sqlCmd_MDB = New OleDb.OleDbCommand(strSQL, MDBConnect1)
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

        '크리스탈 레포트에서 인쇄하는 규칙
        Try
            Dim rptPath As String = Application.StartupPath & "\Reports\Device_Data_RPT.rpt"

            Dim ds As DataSet = New DataSet

            Dim connection As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath + "\TempDB\TempDB_MMPS.mdb" & ";Jet OLEDB:Database Password='dbwlspark'")
            connection.Open()

            Dim strQuery As String = "SELECT `DEVICE_DATA_PRT`.`CUSTOMER_CODE`, `DEVICE_DATA_PRT`.`CUSTOMER_NAME`"
            strQuery += ", `DEVICE_DATA_PRT`.`MODEL_CODE`, `DEVICE_DATA_PRT`.`MODEL_NAME`, `DEVICE_DATA_PRT`.`LINE_NAME`"
            strQuery += ", `DEVICE_DATA_PRT`.`SIDE`, `DEVICE_DATA_PRT`.`MACHINE_NO`, `DEVICE_DATA_PRT`.`FEEDER_NO`"
            strQuery += ", `DEVICE_DATA_PRT`.`PART_NO`, `DEVICE_DATA_PRT`.`SPECIFICATION`, `DEVICE_DATA_PRT`.`DD_NOTE`"
            strQuery += ", `DEVICE_DATA_PRT`.`DD_CODE`, `DEVICE_DATA_PRT`.`MATCH_COUNT`, `DEVICE_DATA_PRT`.`DD_MAIN_NO`"
            strQuery += " FROM   `DEVICE_DATA_PRT` `DEVICE_DATA_PRT`"
            strQuery += " ORDER BY `DEVICE_DATA_PRT`.`MACHINE_NO`, `DEVICE_DATA_PRT`.`FEEDER_NO`"

            'Console.WriteLine(strQuery)

            Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(strQuery, connection)
            da.Fill(ds)

            Dim rDOC As ReportDocument = New ReportDocument

            rDOC.Load(rptPath)
            rDOC.SetDataSource(ds)
            'rDOC.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape '용지방향 설정
            DeviceData_Report.CrystalReportViewer1.ReportSource = Nothing
            DeviceData_Report.CrystalReportViewer1.ReportSource = rDOC

            If MsgBox("인쇄 미리보기를 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.Yes Then
                DeviceData_Report.ShowDialog()
            Else
                rDOC.PrintToPrinter(1, True, 0, 0)
            End If

            connection.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
        End Try

    End Sub

    Private Sub Grid_DeviceData_CellButtonClick(sender As Object, e As RowColEventArgs) Handles Grid_DeviceData.CellButtonClick

        If Grid_DeviceData(e.Row, 0).ToString = "N" Then
            MsgBox("신규 항목의 대치자재는 등록할 수 없습니다." & vbCrLf & "저장 후 다시 시도하여 주십시오.",
                   MsgBoxStyle.Information,
                   form_Msgbox_String)
            Exit Sub
        End If

        Dim sel_machine As Integer = Grid_DeviceData(e.Row, 2)
        Dim sel_feeder As Integer = Grid_DeviceData(e.Row, 3)
        Dim sel_code As String = Grid_DeviceData(e.Row, 1)
        Dim sel_part As String = Grid_DeviceData(e.Row, 5)
        Dim sel_part_maker As String = Grid_DeviceData(e.Row, 4)

        DeviceData_Sub.Tb_machineNo.Text = sel_machine
        DeviceData_Sub.Tb_feederNo.Text = sel_feeder
        DeviceData_Sub.Tb_DDCode.Text = sel_code
        DeviceData_Sub.Tb_mainParts.Text = sel_part
        DeviceData_Sub.ddMainCode = newDeviceData
        DeviceData_Sub.Tb_Customer.Text = Cb_customerName.Text
        DeviceData_Sub.TB_mainParts_Maker.Text = sel_part_maker

        If Not DeviceData_Sub.Visible Then DeviceData_Sub.Show()
        DeviceData_Sub.Focus()

    End Sub

    Dim selFactoryCode As String

    Private Sub Cb_FactoryName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cb_FactoryName.SelectedIndexChanged

        writeReady = False

        DBConnect()

        Dim strSQL As String = "select SUB_CODE from TB_SUB_CODE"
        strSQL += " where SUB_CODE_NAME = '" & Cb_FactoryName.Text & "'"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            selFactoryCode = sqlDR("SUB_CODE")
        Loop
        sqlDR.Close()

        DBClose()

        Cb_workLine.SelectedIndex = -1
        Cb_workSide.SelectedIndex = -1
        Grid_DeviceData.Rows.Count = 1

    End Sub

    Private Sub Cb_FactoryName_DropDown(sender As Object, e As EventArgs) Handles Cb_FactoryName.DropDown


        Dim cb_old_string As String = Cb_FactoryName.Text

        Cb_FactoryName.Items.Clear()

        DBConnect()

        Dim strSQL As String = "select SUB_CODE_NAME from TB_SUB_CODE"
        strSQL += " where MAIN_CODE = 'MC0002' order by SUB_CODE_NAME"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Cb_FactoryName.Items.Add(sqlDR("SUB_CODE_NAME"))
        Loop
        sqlDR.Close()

        DBClose()

        Cb_FactoryName.Text = cb_old_string

    End Sub

    Private Sub Cb_workLine_DropDown(sender As Object, e As EventArgs) Handles Cb_workLine.DropDown

        If Cb_FactoryName.Text = String.Empty Then
            MsgBox("공장을 먼저 선택하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        Dim cb_old_string As String = Cb_workLine.Text

        Cb_workLine.Items.Clear()

        DBConnect()

        Dim strSQL As String = "select LAST_CODE_NAME from TB_LAST_CODE"
        strSQL += " where SUB_CODE = '" & selFactoryCode & "' order by LAST_CODE_NAME"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Cb_workLine.Items.Add(sqlDR("LAST_CODE_NAME"))
        Loop
        sqlDR.Close()

        DBClose()

        Cb_workLine.Text = cb_old_string

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click

        If Not CadData.Visible Then CadData.Show()
        CadData.Focus()

    End Sub

    Private Sub BTN_MX_Feeder_Click(sender As Object, e As EventArgs) Handles BTN_MX_Feeder.Click

        If writeReady = False Then
            MsgBox("조회된 내용이 없습니다." & vbCrLf & "목록 선택 후 조회를 먼저 실행 하십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        Grid_DeviceData.Redraw = False

        'Try
        Dim openFile As New System.Windows.Forms.OpenFileDialog

        openFile.FileName = ""
        openFile.Filter = "MIRAE Program Files (*.prg)|*.prg"
        openFile.ShowDialog()

        If openFile.FileName = "" Then Exit Sub '파일 OPEN 안할때 나가라

        Dim fs As System.IO.FileStream
        Dim sr As System.IO.StreamReader
        fs = System.IO.File.Open(openFile.FileName, IO.FileMode.Open) ' 파일 열기
        sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결

        Dim str As String = String.Empty
        Dim writeOK As Boolean = False
        Dim str2() As String = {}

        Dim machineNO As Integer = 0
        Dim nextMachineNO As Boolean = False

        If IsNumeric(Grid_DeviceData(Grid_DeviceData.Rows.Count - 1, 2)) Then
            machineNO = CInt(Grid_DeviceData(Grid_DeviceData.Rows.Count - 1, 2)) + 1
        Else
            machineNO = 1
        End If

        While sr.Peek() >= 0
            str = sr.ReadLine() ' 한줄씩 읽기
            If str = "@STARTPICKDATA" Then
                writeOK = True
            ElseIf str = "@ENDPICKDATA" Then
                Exit While
            End If

            If writeOK = True Then
                For i = 0 To 10
                    str = str.Replace("  ", " ")
                Next

                If Not Trim(str) = String.Empty Then
                    str2 = Split(str, " ")
                    'Console.WriteLine(str & ",    " & UBound(str2))
                    If UBound(str2) = 14 Then
                        Dim partInfo As String = loadMakerSpec(str2(3))
                        If CInt(str2(2)) > 50 And nextMachineNO = False Then
                            nextMachineNO = True
                            machineNO += 1
                        End If
                        Dim insert_string As String = "N" & vbTab &
                          DeviceDataCode() & vbTab &
                          machineNO & vbTab &
                          str2(2) & vbTab &
                          partInfo.Split("@")(0) & vbTab &
                          str2(3) & vbTab &
                          partInfo.Split("@")(1) & vbTab &
                          String.Empty & vbTab &
                          String.Empty & vbTab &
                          String.Empty
                        Grid_DeviceData.AddItem(insert_string)
                        Grid_DeviceData.Rows(Grid_DeviceData.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
                    End If
                End If
            End If
        End While
        sr.Close()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
        'End Try

        Grid_DeviceData.AutoSizeCols()
        Grid_DeviceData.AutoSizeRows(1, 0, Grid_DeviceData.Rows.Count - 1, Grid_DeviceData.Cols.Count - 1, 0, AutoSizeFlags.None)

        Grid_DeviceData.Redraw = True

        MsgBox("MIRAE Feeder List 불러오기를 완료 하였습니다.", MsgBoxStyle.Information, form_Msgbox_String)

    End Sub

    Private Function loadMakerSpec(ByVal partNo As String)

        loadMakerSpec = String.Empty

        DBConnect()

        Dim strSQL As String = "select PART_MAKER, PART_SPECIFICATION from TB_DEVICE_DATA where PART_NO = '" & partNo & "' limit 1;"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            loadMakerSpec = sqlDR("PART_MAKER") & "@" & sqlDR("PART_SPECIFICATION")
        Loop
        sqlDR.Close()

        DBClose()

        If loadMakerSpec = String.Empty Then loadMakerSpec = "@"

        Return loadMakerSpec

    End Function

    Private Sub BTN_SM_Feeder2_Click(sender As Object, e As EventArgs) Handles BTN_SM_Feeder2.Click

        If writeReady = False Then
            MsgBox("조회된 내용이 없습니다." & vbCrLf & "목록 선택 후 조회를 먼저 실행 하십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        Grid_DeviceData.Redraw = False

        'Try
        Dim openFile As New System.Windows.Forms.OpenFileDialog

        openFile.FileName = ""
        openFile.Filter = "SAMSUNG Feeder List(*.csv)|*.csv"
        openFile.ShowDialog()

        If openFile.FileName = "" Then Exit Sub '파일 OPEN 안할때 나가라

        Dim fs As System.IO.FileStream
        Dim sr As System.IO.StreamReader
        fs = System.IO.File.Open(openFile.FileName, IO.FileMode.Open) ' 파일 열기
        sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결

        Dim str As String = String.Empty
        Dim writeOK As Boolean = False
        Dim str2() As String = {}

        Dim machineNO As Integer = 0
        Dim nextMachineNO As Boolean = False

        If IsNumeric(Grid_DeviceData(Grid_DeviceData.Rows.Count - 1, 2)) Then
            machineNO = CInt(Grid_DeviceData(Grid_DeviceData.Rows.Count - 1, 2)) + 1
        Else
            machineNO = 1
        End If

        While sr.Peek() >= 0
            str = sr.ReadLine() ' 한줄씩 읽기
            If Not str.Substring(0, 7) = "Machine" Then
                writeOK = True
            End If

            If writeOK = True Then
                If Not Trim(str) = String.Empty Then
                    str2 = Split(str, ",")

                    machineNO = CDbl(str2(0).Substring(1, 1)) * 2 - 1
                    If str2(1).Substring(0, 1) = "R" Then
                        machineNO += 1
                    End If
                    Dim partInfo As String = loadMakerSpec(str2(3))
                    Dim insert_string As String = "N" & vbTab &
                          DeviceDataCode() & vbTab &
                          machineNO & vbTab &
                          str2(1).Split("@")(1) & vbTab &
                          partInfo.Split("@")(0) & vbTab &
                          str2(3) & vbTab &
                          partInfo.Split("@")(1) & vbTab &
                          String.Empty & vbTab &
                          String.Empty & vbTab &
                          String.Empty
                    Grid_DeviceData.AddItem(insert_string)
                    Grid_DeviceData.Rows(Grid_DeviceData.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
                End If
            End If
        End While
        sr.Close()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
        'End Try

        Grid_DeviceData.AutoSizeCols()
        Grid_DeviceData.AutoSizeRows(1, 0, Grid_DeviceData.Rows.Count - 1, Grid_DeviceData.Cols.Count - 1, 0, AutoSizeFlags.None)

        Grid_DeviceData.Redraw = True

        MsgBox("SAMSUNG(HANWHA) Feeder List(*.csv) 불러오기를 완료 하였습니다.", MsgBoxStyle.Information, form_Msgbox_String)

    End Sub

    Private Sub TB_MakerCheck_Click(sender As Object, e As EventArgs) Handles TB_MakerCheck.Click

        If Grid_DeviceData.Rows.Count = 1 Then
            MsgBox("조회를 먼저하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        DBConnect()

        For i = 1 To Grid_DeviceData.Rows.Count - 1
            If Grid_DeviceData(i, 4).Equals("") Then
                Dim strSQL As String = "select PART_MAKER from TB_DEVICE_DATA where PART_NO = '" & Grid_DeviceData(i, 5) & "' limit 1"

                Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
                Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

                Do While sqlDR.Read
                    If Not sqlDR("PART_MAKER").Equals("") Then
                        Grid_DeviceData(i, 0) = "M"
                        Grid_DeviceData.Rows(i).StyleNew.ForeColor = Color.Red
                        Grid_DeviceData(i, 4) = sqlDR("PART_MAKER")
                    End If
                Loop
                sqlDR.Close()
            End If
        Next

        DBClose()

        Grid_DeviceData.AutoSizeCols()
        Grid_DeviceData.Redraw = True

        MsgBox("자재 제조사 조회가 완료 되었습니다." & vbCrLf & "변경사항 확인 후 저장하여 주십시오.",
               MsgBoxStyle.Information,
               form_Msgbox_String)

    End Sub

    Private Sub Grid_DeviceData_AfterFilter(sender As Object, e As EventArgs) Handles Grid_DeviceData.AfterFilter

    End Sub

    Private Sub DeviceData_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        Mainform.form_name = Me

    End Sub

    Private Sub Grid_DeviceData_KeyDown(sender As Object, e As KeyEventArgs) Handles Grid_DeviceData.KeyDown

        If writeReady = False Then
            MsgBox("조회된 내용이 없습니다." & vbCrLf & "목록 선택 후 조회를 먼저 실행 하십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If e.KeyCode = 112 Then
            BTN_newLine_Click(Nothing, Nothing)
        End If

    End Sub
End Class


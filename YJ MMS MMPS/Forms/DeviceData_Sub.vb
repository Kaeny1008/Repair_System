Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class DeviceData_Sub

    Dim form_Msgbox_String As String = "대치자재 관리"
    Public ddMainCode As String

    Private Sub DeviceData_Sub_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GridSetting()

        Data_Load()

    End Sub

    Private Sub GridSetting()

        With Grid_DeviceData
            .AllowEditing = True
            .AllowFiltering = True
            .AutoClipboard = True
            .AllowSorting = AllowSortingEnum.None
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 5
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            Grid_DeviceData(0, 0) = "No"
            Grid_DeviceData(0, 1) = "Code"
            Grid_DeviceData(0, 2) = "Maker"
            Grid_DeviceData(0, 3) = "Part No."
            Grid_DeviceData(0, 4) = "비고"
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 1).TextAlign = TextAlignEnum.LeftCenter
            .Cols(2).ComboList = Maker_List()
            .AutoSizeCols()
            .ExtendLastCol = True
        End With

    End Sub

    Private Sub Data_Load()

        DBConnect()

        Grid_DeviceData.Redraw = False
        Grid_DeviceData.Rows.Count = 1

        Dim strSQL As String = String.Empty
        If Tb_Customer.Text = "J산업" Or
            Tb_Customer.Text = "전유산업" Or
            Tb_Customer.Text = "L-Tech" Or
            Tb_Customer.Text = "allRadio" Or
            Tb_Customer.Text = "TopRun" Or
            Tb_Customer.Text = "덕일전자" Or
            Tb_Customer.Text.Contains("CI Digital") Or
            Tb_Customer.Text = "Mangoslab" Then
            strSQL = "select DS_CODE, PART_MAKER, PART_NO, DS_NOTE from TB_DEVICE_DATA_SUB2"
            strSQL += " where MAIN_PART_NO = '" & Tb_mainParts.Text & "' and CUSTOMER = '" & Tb_Customer.Text & "' order by PART_NO"
        Else
            strSQL = "Select DS_CODE, PART_MAKER, PART_NO, DS_NOTE from TB_DEVICE_DATA_SUB"
            strSQL += " where DD_CODE = '" & Tb_DDCode.Text & "' order by PART_NO"
        End If

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_string As String = Grid_DeviceData.Rows.Count & vbTab &
                                              sqlDR("DS_CODE") & vbTab &
                                              sqlDR("PART_MAKER") & vbTab &
                                              sqlDR("PART_NO") & vbTab &
                                              sqlDR("DS_NOTE")
            Grid_DeviceData.AddItem(insert_string)
        Loop
        sqlDR.Close()

        If Grid_DeviceData.Rows.Count = 1 Then
            Grid_DeviceData.AddItem("N" & vbTab & NEW_Code())
            Grid_DeviceData.Rows(1).StyleNew.ForeColor = Color.Blue
        End If

        Grid_DeviceData.AutoSizeCols()
        Grid_DeviceData.Redraw = True

        DBClose()

    End Sub

    Private Sub Grid_DeviceData_MouseClick(sender As Object, e As MouseEventArgs) Handles Grid_DeviceData.MouseClick

        Dim row As Integer = Grid_DeviceData.MouseRow

        If e.Button = MouseButtons.Right And row > -1 Then
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

    End Sub

    Private Function NEW_Code() As String

        Dim newCode As String = String.Empty

        DBConnect()

        Dim strSQL As String = String.Empty
        If Tb_Customer.Text = "J산업" Or
            Tb_Customer.Text = "전유산업" Or
            Tb_Customer.Text = "TopRun" Or
            Tb_Customer.Text = "덕일전자" Or
            Tb_Customer.Text = "CI Digital" Or
            Tb_Customer.Text = "L-Tech" Or
            Tb_Customer.Text.Contains("CI Digital") Or
            Tb_Customer.Text = "Mangoslab" Then
            strSQL = "select DS_CODE from TB_DEVICE_DATA_SUB2 order by DS_CODE desc limit 1"
        Else
            strSQL = "select DS_CODE from TB_DEVICE_DATA_SUB order by DS_CODE desc limit 1"
        End If

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            newCode = sqlDR("DS_CODE")
        Loop
        sqlDR.Close()

        DBClose()

        If newCode = String.Empty Then
            newCode = "DS00000001"
        Else
            newCode = "DS" & Format(CInt(newCode.Substring(2, 8)) + 1, "00000000")
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

    Private Sub BTN_newLine_Click(sender As Object, e As EventArgs) Handles BTN_newLine.Click

        Dim insertRow As Integer = 0
        If Grid_DeviceData.Row < 1 Then
            insertRow = 1
        Else
            insertRow = Grid_DeviceData.Row + 1
        End If

        Grid_DeviceData.Redraw = False

        Grid_DeviceData.AddItem("N" & vbTab & NEW_Code(), insertRow)
        Grid_DeviceData.Rows(insertRow).StyleNew.ForeColor = Color.Blue

        Grid_DeviceData.AutoSizeCols()

        Grid_DeviceData.Redraw = True

    End Sub

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

    Private Sub BTN_Save_Click(sender As Object, e As EventArgs) Handles BTN_Save.Click

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

            If Tb_Customer.Text = "J산업" Or
                Tb_Customer.Text = "전유산업" Or
                Tb_Customer.Text = "allRadio" Or
                Tb_Customer.Text = "L-Tech" Or
                Tb_Customer.Text = "덕일전자" Or
                Tb_Customer.Text.Contains("CI Digital") Or
                Tb_Customer.Text = "TopRun" Or
                Tb_Customer.Text = "Mangoslab" Then
                For i = 1 To Grid_DeviceData.Rows.Count - 1
                    If Grid_DeviceData(i, 0).ToString = "N" Then
                        strSQL += "Insert Into TB_DEVICE_DATA_SUB2(DS_CODE, CUSTOMER, MAIN_PART_MAKER, MAIN_PART_NO, PART_MAKER, PART_NO, DS_NOTE) values"
                        strSQL += "('" & Grid_DeviceData(i, 1) & "'"
                        strSQL += ",'" & Tb_Customer.Text & "'"
                        strSQL += ",'" & TB_mainParts_Maker.Text & "'"
                        strSQL += ",'" & Tb_mainParts.Text & "'"
                        strSQL += ",'" & Grid_DeviceData(i, 2) & "'"
                        strSQL += ",'" & Grid_DeviceData(i, 3) & "'"
                        strSQL += ",'" & Grid_DeviceData(i, 4) & "');"
                    ElseIf Grid_DeviceData(i, 0).ToString = "M" Then
                        strSQL += "update TB_DEVICE_DATA_SUB2 set"
                        strSQL += " PART_MAKER = '" & Grid_DeviceData(i, 2) & "'"
                        strSQL += ", PART_NO = '" & Grid_DeviceData(i, 3) & "'"
                        strSQL += ", DS_NOTE = '" & Grid_DeviceData(i, 4) & "'"
                        strSQL += " where DS_CODE = '" & Grid_DeviceData(i, 1) & "'"
                        strSQL += " and CUSTOMER = '" & Tb_Customer.Text & "';"
                    ElseIf Grid_DeviceData(i, 0).ToString = "D" Then
                        strSQL += "delete from TB_DEVICE_DATA_SUB2"
                        strSQL += " where DS_CODE = '" & Grid_DeviceData(i, 1) & "'"
                        strSQL += " and CUSTOMER = '" & Tb_Customer.Text & "';"
                    End If
                Next
            Else
                For i = 1 To Grid_DeviceData.Rows.Count - 1
                    If Grid_DeviceData(i, 0).ToString = "N" Then
                        '동일한 DD_MAIN_CODE내의 같은 PARTS No를 가지고 있다면 동일하게 적용하여 업데이트
                        strSQL += "Insert Into TB_DEVICE_DATA_SUB(DD_CODE, DS_CODE, PART_MAKER, PART_NO, DS_NOTE)"
                        strSQL += " select DD_CODE"
                        strSQL += ",'" & Grid_DeviceData(i, 1) & "'"
                        strSQL += ",'" & Grid_DeviceData(i, 2) & "'"
                        strSQL += ",'" & Grid_DeviceData(i, 3) & "'"
                        strSQL += ",'" & Grid_DeviceData(i, 4) & "'"
                        strSQL += " from TB_DEVICE_DATA"
                        strSQL += " where DD_MAIN_NO = '" & ddMainCode & "'"
                        strSQL += " and PART_NO = '" & Tb_mainParts.Text & "';"
                    ElseIf Grid_DeviceData(i, 0).ToString = "M" Then
                        strSQL += "update TB_DEVICE_DATA_SUB set"
                        strSQL += " PART_MAKER = '" & Grid_DeviceData(i, 2) & "'"
                        strSQL += ", PART_NO = '" & Grid_DeviceData(i, 3) & "'"
                        strSQL += ", DS_NOTE = '" & Grid_DeviceData(i, 4) & "'"
                        strSQL += " where DS_CODE = '" & Grid_DeviceData(i, 1) & "';"
                    ElseIf Grid_DeviceData(i, 0).ToString = "D" Then
                        strSQL += "delete from TB_DEVICE_DATA_SUB"
                        strSQL += " where DS_CODE = '" & Grid_DeviceData(i, 1) & "';"
                    End If
                Next
            End If

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

        Data_Load()

    End Sub

    Private Sub Grid_DeviceData_AfterEdit(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles Grid_DeviceData.AfterEdit

        If before_text = Grid_DeviceData(e.Row, e.Col) Then Exit Sub

        If IsNumeric(Grid_DeviceData(e.Row, 0)) Then
            Grid_DeviceData(e.Row, 0) = "M"
            Grid_DeviceData.Rows(e.Row).StyleNew.ForeColor = Color.Red
        End If

        Grid_DeviceData.AutoSizeCols()

    End Sub

    Dim before_text As String

    Private Sub Grid_DeviceData_BeforeEdit(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles Grid_DeviceData.BeforeEdit

        before_text = Grid_DeviceData(e.Row, e.Col)

    End Sub

    Private Sub Grid_DeviceData_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Grid_DeviceData.KeyDown

        If Grid_DeviceData(Grid_DeviceData.Row, 0).ToString = "D" Then
            e.SuppressKeyPress = True
        End If

    End Sub

    Private Sub DeviceData_Sub_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        DeviceData.BTN_Search_Click(Nothing, Nothing)

    End Sub

    Private Sub Grid_DeviceData_ComboDropDown(sender As Object, e As RowColEventArgs) Handles Grid_DeviceData.ComboDropDown

        If e.Col = 2 Then
            Grid_DeviceData.Cols(2).ComboList = Maker_List()
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

    Private Sub Grid_DeviceData_RowColChange(sender As Object, e As EventArgs) Handles Grid_DeviceData.RowColChange

        Select Case Grid_DeviceData.Col
            Case 1
                Grid_DeviceData.AllowEditing = False
            Case 2, 3, 4
                Grid_DeviceData.AllowEditing = True
        End Select

    End Sub
End Class
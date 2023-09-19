'############################################################################################################
'############################################################################################################

'##### 폼제목   : OMS 파일 기초정보
'##### 작성일자 : 2023-07-10
'##### 수정일자 : 2023-07-10
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : OMS에 등록시킬 파일의 테이블 정의
'                 각 종류별 별도관리하여 저장될 수 있게

'############################################################################################################
'############################################################################################################

Imports System.Threading
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_oms_file_data

    Dim form_msgbox_string As String = "OMS 파일 기초정보"

    Private Sub OMS_FILE_INFO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        GridSetting()

    End Sub

    Private Sub GridSetting()

        With grid_Data_List
            .AllowEditing = True
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 9
            .Cols.Fixed = 1
            .Rows.Count = 2
            .Rows.Fixed = 1
            grid_Data_List(0, 0) = "No"
            grid_Data_List(0, 1) = "Field ID"
            grid_Data_List(0, 2) = "Null"
            grid_Data_List(0, 3) = "Field Name"
            grid_Data_List(0, 4) = "Type"
            grid_Data_List(0, 5) = "Data Length"
            grid_Data_List(0, 6) = "Accumulation"
            grid_Data_List(0, 7) = "Description"
            grid_Data_List(0, 8) = "Data Code"
            .Cols(8).Visible = False
            '.Cols(1).ComboList = "대표|대체"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 2).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            grid_Data_List(1, 0) = "N"
            .Rows(1).StyleNew.ForeColor = Color.Blue
            .AutoSizeCols()
            .ShowCursor = True
        End With

    End Sub

    Private Sub TB_File_Type_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_File_Type.SelectedIndexChanged

        If TB_File_Type.Text = "CMWIPINF" Then
            TB_File_Desc.Text = "재공"
            TB_IF_Name.Text = "CMWIPINF_XXYYYYMMDDHH.ff"
            TB_Frequency.Text = "Every 2 hours (from 00:10)"
        ElseIf TB_File_Type.Text = "CMHALFIN" Then
            TB_File_Desc.Text = "반제품입고"
            TB_IF_Name.Text = "CMHALFIN_XXYYYYMMDDHH.ff"
            TB_Frequency.Text = "Immediately"
        End If

        grid_Data_List.Clear()
        GridSetting()

    End Sub

    Dim before_data As String

    Private Sub GRID_Data_List_BeforeEdit(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles grid_Data_List.BeforeEdit

        before_data = grid_Data_List(e.Row, e.Col)

    End Sub

    Private Sub GRID_Data_List_AfterEdit(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles grid_Data_List.AfterEdit

        If before_data = grid_Data_List(e.Row, e.Col) Then Exit Sub

        If IsNumeric(grid_Data_List(e.Row, 0)) Then
            grid_Data_List(e.Row, 0) = "M"
            grid_Data_List.Rows(e.Row).StyleNew.ForeColor = Color.Red
        End If

        grid_Data_List.Redraw = False

        Dim data_code As Double = Int(Format(Now, "yyyyMMddHHmmss")) - 1

        For i = 1 To grid_Data_List.Rows.Count - 1
            If grid_Data_List(i, 0).ToString = "N" Then
                data_code += 1
                grid_Data_List(i, 8) = data_code
            End If
        Next
        grid_Data_List.AutoSizeCols()

        grid_Data_List.Redraw = True

    End Sub

    Private Sub GRID_Data_List_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grid_Data_List.KeyDown

        If e.KeyCode = Keys.V AndAlso e.Modifiers = Keys.Control Then 'Ctrl + V
            Dim insert_Count As Integer = UBound(Clipboard.GetText.Split(Chr(10))) - 1
            Dim start_Row As Integer = grid_Data_List.Row
            Dim total_Row As Integer = grid_Data_List.Rows.Count
            Dim empty_Row As Integer = total_Row - start_Row

            For i = 1 To insert_Count - empty_Row + 1
                grid_Data_List.AddItem("N", grid_Data_List.Row + 1)
            Next

            GRID_Num_Sort()
        Else
            If grid_Data_List(grid_Data_List.Row, 0).ToString = "D" Then
                e.SuppressKeyPress = True
            End If
        End If

    End Sub

    Private Sub GRID_Data_List_RowColChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid_Data_List.RowColChange

        'Dim sel_Row As Integer = GRID_Data_List.Row

        'If sel_Row < 1 Then Exit Sub

        'If GRID_Data_List(sel_Row, 0).ToString = "D" Then
        '    GRID_Data_List.AllowEditing = False
        'Else
        '    GRID_Data_List.AllowEditing = True
        'End If

    End Sub

    Private Sub grid_Data_List_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grid_Data_List.MouseClick

        If e.Button = Windows.Forms.MouseButtons.Right And grid_Data_List.MouseRow > 0 Then
            grid_Data_List.Row = grid_Data_List.MouseRow
            GRID_Menu.Show(grid_Data_List, New Point(e.X, e.Y))
        End If

    End Sub

    Private Sub btn_Row_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Row_Add.Click

        grid_Data_List.AddItem("N", grid_Data_List.Row + 1)

        GRID_Num_Sort()

    End Sub

    Private Sub btn_Row_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Row_Delete.Click

        Dim sel_Row As Integer = grid_Data_List.Row

        If grid_Data_List(sel_Row, 0).ToString = "N" Then
            grid_Data_List.RemoveItem(sel_Row)
        Else
            grid_Data_List(sel_Row, 0) = "D"
        End If

        GRID_Num_Sort()

    End Sub

    Private Sub GRID_Num_Sort()

        grid_Data_List.Redraw = False

        For i = 1 To grid_Data_List.Rows.Count - 1
            If IsNumeric(grid_Data_List(i, 0).ToString) Then
                grid_Data_List(i, 0) = i
            Else
                If grid_Data_List(i, 0).ToString = "N" Then
                    grid_Data_List.Rows(i).StyleNew.ForeColor = Color.Blue
                ElseIf grid_Data_List(i, 0).ToString = "D" Then
                    grid_Data_List.Rows(i).StyleNew.ForeColor = Color.Gray
                End If
            End If
        Next

        grid_Data_List.AutoSizeCols()

        grid_Data_List.Redraw = True

    End Sub

    Private Sub btn_SAVE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_SAVE.Click, btn_Menu_Save.Click

        If TB_File_Type.Text = String.Empty Then
            MsgBox("File Type이 선택되지 않았습니다.", MsgBoxStyle.Information, form_msgbox_string)
            Exit Sub
        End If

        For i = 1 To grid_Data_List.Rows.Count - 1
            For j = 1 To grid_Data_List.Cols.Count - 1
                If CStr(grid_Data_List(i, j)) = String.Empty And Not j = 7 Then
                    MsgBox("입력되지않은 항목이 존재 합니다." & vbCrLf &
                           "Description 제외 모든 항목이 입력되어야 합니다.", MsgBoxStyle.Information, form_msgbox_string)
                    Exit Sub
                End If
            Next
        Next

        If MsgBox("저장 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Saving...")

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try

            '##### FILE DATA를 저장한다. 시작 #####

            For i = 1 To grid_Data_List.Rows.Count - 1
                If grid_Data_List(i, 0).ToString = "N" Then
                    '현재행이 신규 입력행일 경우
                    strSQL += "insert into setting_oms_file_data(DATA_CODE, FILE_TYPE, DATA_SORT, FIELD_ID, FIELD_NULL"
                    strSQL += ", FIELD_NAME, FIELD_TYPE, DATA_LENGTH, ACCUMULATION, FIELD_DESC, WRITER) values"
                    strSQL += "('" & grid_Data_List(i, 8) & "'"
                    strSQL += ",'" & TB_File_Type.Text & "'"
                    strSQL += ",'" & i & "'"
                    strSQL += ",'" & grid_Data_List(i, 1) & "'"
                    strSQL += ",'" & grid_Data_List(i, 2) & "'"
                    strSQL += ",'" & grid_Data_List(i, 3) & "'"
                    strSQL += ",'" & grid_Data_List(i, 4) & "'"
                    strSQL += ",'" & grid_Data_List(i, 5) & "'"
                    strSQL += ",'" & grid_Data_List(i, 6) & "'"
                    strSQL += ",'" & grid_Data_List(i, 7) & "'"
                    strSQL += ",'" & loginID & "');"
                ElseIf grid_Data_List(i, 0).ToString = "M" Then
                    '현재행이 수정된 경우
                    strSQL += "update setting_oms_file_data set DATA_SORT = '" & i & "'"
                    strSQL += ", FIELD_ID = '" & grid_Data_List(i, 1) & "'"
                    strSQL += ", FIELD_NULL = '" & grid_Data_List(i, 2) & "'"
                    strSQL += ", FIELD_NAME = '" & grid_Data_List(i, 3) & "'"
                    strSQL += ", FIELD_TYPE = '" & grid_Data_List(i, 4) & "'"
                    strSQL += ", DATA_LENGTH = '" & grid_Data_List(i, 5) & "'"
                    strSQL += ", ACCUMULATION = '" & grid_Data_List(i, 6) & "'"
                    strSQL += ", FIELD_DESC = '" & grid_Data_List(i, 7) & "'"
                    strSQL += ", WRITER = '" & loginID & "'"
                    strSQL += " where DATA_CODE = '" & grid_Data_List(i, 8) & "'"
                ElseIf grid_Data_List(i, 0).ToString = "D" Then
                    '현재행을 삭제 할 경우
                    strSQL += "delete from setting_oms_file_data where DATA_CODE = '" & grid_Data_List(i, 8) & "';"
                End If
            Next

            '##### FILE DATA를 저장한다. 종료 #####
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

        th_LoadingWindow.Abort()
        Thread.Sleep(100)
        MsgBox("저장완료.", MsgBoxStyle.Information, form_msgbox_string)

        btn_Search_Click(Nothing, Nothing)

    End Sub

    Private Sub btn_Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Search.Click

        If TB_File_Type.Text = String.Empty Then
            MsgBox("File Type이 선택되지 않았습니다.", MsgBoxStyle.Information, form_msgbox_string)
            Exit Sub
        End If

        thread_LoadingFormStart()

        grid_Data_List.Redraw = False
        grid_Data_List.Rows.Count = 1

        DBConnect()

        'strSQL += "insert into setting_oms_file_data(DATA_CODE, FILE_TYPE, DATA_SORT, FIELD_ID, FIELD_NULL"
        'strSQL += ", FIELD_NAME, FIELD_TYPE, DATA_LENGTH, ACCUMULATION, FIELD_DESC) values"

        Dim strSQL As String = "select data_sort, field_id, field_null, field_name, field_type, data_length, accumulation, field_desc, data_code"
        strSQL += " from setting_oms_file_data"
        strSQL += " where file_type = '" & TB_File_Type.Text & "'"
        strSQL += " order by data_sort"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = grid_Data_List.Rows.Count & vbTab &
                                          sqlDR("FIELD_ID") & vbTab &
                                          sqlDR("FIELD_NULL") & vbTab &
                                          sqlDR("FIELD_NAME") & vbTab &
                                          sqlDR("FIELD_TYPE") & vbTab &
                                          sqlDR("DATA_LENGTH") & vbTab &
                                          sqlDR("ACCUMULATION") & vbTab &
                                          sqlDR("FIELD_DESC") & vbTab &
                                          sqlDR("DATA_CODE")
            grid_Data_List.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        If grid_Data_List.Rows.Count = 1 Then
            grid_Data_List.AddItem("N")
            GRID_Num_Sort()
        End If

        grid_Data_List.AutoSizeCols()
        grid_Data_List.Redraw = True

        th_LoadingWindow.Abort()

    End Sub

    Private Sub COL_VIEW_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles COL_VIEW.MouseDown

        grid_Data_List.Cols(8).Visible = True

    End Sub

    Private Sub COL_VIEW_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles COL_VIEW.MouseUp

        grid_Data_List.Cols(8).Visible = False

    End Sub
End Class
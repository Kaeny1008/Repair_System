'############################################################################################################
'############################################################################################################

'##### 폼제목   : Bucket별 PMIC/RCD 조합 및 Vendor
'##### 작성일자 : 2023-07-10
'##### 수정일자 : 2023-07-10
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : Lot Card의 Product Code 17, 18자리의 Customer Code를 확인하여 해당 PMIC, RCD 자재를 찾는다.
'                 

'############################################################################################################
'############################################################################################################

Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class frm_Bucket_RCD_PMIC

    Dim form_msgbox_string As String = "Bucket별 PMIC/RCD 조합 및 Vendor"

    Private Sub frm_Bucket_RCD_PMIC_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        GridSetting()

    End Sub

    Private Sub GridSetting()

        With GRID_Data_List
            .AllowEditing = True
            .AllowFiltering = True
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 12
            .Cols.Fixed = 1
            .Rows.Count = 2
            .Rows.Fixed = 2
            .AutoClipboard = True
            .AllowFreezing = AllowFreezingEnum.Both
            .AllowMerging = AllowMergingEnum.RestrictAll
            .AllowMergingFixed = AllowMergingEnum.FixedOnly
            For i = 0 To .Cols.Count - 1
                .Cols(i).AllowMerging = True
            Next
            .Rows(0).AllowMerging = True
            .Rows(1).AllowMerging = True
            Dim Mrng As CellRange = .GetCellRange(0, 0, 1, 0)
            Mrng.Data = "No"
            Mrng = .GetCellRange(0, 1, 1, 1)
            Mrng.Data = "Customer Code"
            Mrng = .GetCellRange(0, 2, 0, 5)
            Mrng.Data = "PMIC"
            GRID_Data_List(1, 2) = "Vendor"
            GRID_Data_List(1, 3) = "Part No."
            GRID_Data_List(1, 4) = "Top Marking"
            GRID_Data_List(1, 5) = "Vendor Part No."
            Mrng = .GetCellRange(0, 6, 0, 9)
            Mrng.Data = "RCD"
            GRID_Data_List(1, 6) = "Vendor"
            GRID_Data_List(1, 7) = "Part No."
            GRID_Data_List(1, 8) = "Top Marking"
            GRID_Data_List(1, 9) = "Vendor Part No."
            Mrng = .GetCellRange(0, 10, 1, 10)
            Mrng.Data = "비고"
            Mrng = .GetCellRange(0, 11, 1, 11)
            Mrng.Data = "tb_unique"
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .AutoSizeCols()
            .ExtendLastCol = True
            .Cols(11).Visible = False
        End With

    End Sub

    Private Sub grid_Data_List_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GRID_Data_List.MouseClick

        If e.Button = Windows.Forms.MouseButtons.Right And GRID_Data_List.MouseRow > 0 Then
            GRID_Data_List.Row = GRID_Data_List.MouseRow
            cms_Menu.Show(GRID_Data_List, New Point(e.X, e.Y))
        End If

    End Sub

    Private Sub btn_Row_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Row_Add.Click

        GRID_Data_List.AddItem("N", GRID_Data_List.Row + 1)

        GRID_Num_Sort()

    End Sub

    Private Sub btn_Row_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Row_Delete.Click

        Dim sel_Row As Integer = GRID_Data_List.Row

        If GRID_Data_List(sel_Row, 0).ToString = "N" Then
            GRID_Data_List.RemoveItem(sel_Row)
        Else
            GRID_Data_List(sel_Row, 0) = "D"
        End If

        GRID_Num_Sort()

    End Sub

    Private Sub GRID_Num_Sort()

        GRID_Data_List.Redraw = False

        For i = 1 To GRID_Data_List.Rows.Count - 1
            If IsNumeric(GRID_Data_List(i, 0).ToString) Then
                GRID_Data_List(i, 0) = i
            Else
                If GRID_Data_List(i, 0).ToString = "N" Then
                    GRID_Data_List.Rows(i).StyleNew.ForeColor = Color.Blue
                ElseIf GRID_Data_List(i, 0).ToString = "D" Then
                    GRID_Data_List.Rows(i).StyleNew.ForeColor = Color.Gray
                End If
            End If
        Next

        GRID_Data_List.AutoSizeCols()

        GRID_Data_List.Redraw = True

    End Sub

    Private Sub btn_SAVE_Click(sender As Object, e As EventArgs) Handles btn_SAVE.Click, btn_Menu_Save.Click

        For i = 2 To GRID_Data_List.Rows.Count - 1
            For j = 1 To GRID_Data_List.Cols.Count - 1
                If CStr(GRID_Data_List(i, j)) = String.Empty And Not j = 5 And Not j = 9 And Not j = 10 Then
                    MsgBox("입력되지않은 항목이 존재 합니다." & vbCrLf &
                           "Vendor Part No. 및 비고 제외 모든 항목이 입력되어야 합니다." & vbCrLf &
                           "행 : " & i - 1 & "열 : " & j, MsgBoxStyle.Information, form_msgbox_string)
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
        Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

        sqlTran = dbConnection1.BeginTransaction

        Try

            '##### FILE DATA를 저장한다. 시작 #####

            For i = 1 To GRID_Data_List.Rows.Count - 1
                If GRID_Data_List(i, 0).ToString = "N" Then
                    '현재행이 신규 입력행일 경우
                    strSQL += "insert into setting_bucket_combination(customer_code, pmic_vendor, pmic_partno, pmic_top_marking, pmic_org_partno"
                    strSQL += ", rcd_vendor, rcd_partno, rcd_top_marking, rcd_org_partno, note, writer, write_date) values"
                    strSQL += "('" & GRID_Data_List(i, 1) & "'"
                    strSQL += ",'" & GRID_Data_List(i, 2) & "'"
                    strSQL += ",'" & GRID_Data_List(i, 3) & "'"
                    strSQL += ",'" & GRID_Data_List(i, 4) & "'"
                    strSQL += ",'" & GRID_Data_List(i, 5) & "'"
                    strSQL += ",'" & GRID_Data_List(i, 6) & "'"
                    strSQL += ",'" & GRID_Data_List(i, 7) & "'"
                    strSQL += ",'" & GRID_Data_List(i, 8) & "'"
                    strSQL += ",'" & GRID_Data_List(i, 9) & "'"
                    strSQL += ",'" & GRID_Data_List(i, 10) & "'"
                    strSQL += ",'" & loginID & "'"
                    strSQL += ",'" & write_date & "');"
                ElseIf GRID_Data_List(i, 0).ToString = "M" Then
                    '현재행이 수정된 경우
                    strSQL += "update setting_bucket_combination set customer_code = '" & GRID_Data_List(i, 1) & "'"
                    strSQL += ", pmic_vendor = '" & GRID_Data_List(i, 2) & "'"
                    strSQL += ", pmic_partno = '" & GRID_Data_List(i, 3) & "'"
                    strSQL += ", pmic_top_marking = '" & GRID_Data_List(i, 4) & "'"
                    strSQL += ", pmic_org_partno = '" & GRID_Data_List(i, 5) & "'"
                    strSQL += ", rcd_vendor = '" & GRID_Data_List(i, 6) & "'"
                    strSQL += ", rcd_partno = '" & GRID_Data_List(i, 7) & "'"
                    strSQL += ", rcd_top_marking = '" & GRID_Data_List(i, 8) & "'"
                    strSQL += ", rcd_org_partno = '" & GRID_Data_List(i, 9) & "'"
                    strSQL += ", note = '" & GRID_Data_List(i, 10) & "'"
                    strSQL += ", writer = '" & loginID & "'"
                    strSQL += ", write_date = '" & write_date & "'"
                    strSQL += " where tb_unique = '" & GRID_Data_List(i, 11) & "';"
                ElseIf GRID_Data_List(i, 0).ToString = "D" Then
                    '현재행을 삭제 할 경우
                    strSQL += "delete from setting_bucket_combination where customer_code = '" & GRID_Data_List(i, 1) & "';"
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

        'MsgBox("저장완료.", MsgBoxStyle.Information, form_msgbox_string)

        thread_LoadingFormEnd

        btn_Search_Click(Nothing, Nothing)

    End Sub

    Private Sub GRID_Data_List_AfterEdit(sender As Object, e As RowColEventArgs) Handles GRID_Data_List.AfterEdit

        If IsNumeric(GRID_Data_List(e.Row, 0)) Then
            GRID_Data_List(e.Row, 0) = "M"
            GRID_Data_List.Rows(e.Row).StyleNew.ForeColor = Color.Red
        End If

        GRID_Data_List.Redraw = False
        GRID_Data_List.AutoSizeCols()
        GRID_Data_List.Redraw = True

    End Sub

    Private Sub btn_Search_Click(sender As Object, e As EventArgs) Handles btn_Search.Click

        thread_LoadingFormStart()

        GRID_Data_List.Redraw = False
        GRID_Data_List.ClearFilter()
        GRID_Data_List.Rows.Count = 2

        DBConnect()

        Dim strSQL As String = "select customer_code, pmic_vendor, pmic_partno, pmic_top_marking, pmic_org_partno"
        strSQL += ", rcd_vendor, rcd_partno, rcd_top_marking, rcd_org_partno, note, tb_unique"
        strSQL += " from setting_bucket_combination"
        strSQL += " order by customer_code"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert_String As String = GRID_Data_List.Rows.Count - 1 & vbTab &
                                          sqlDR("customer_code") & vbTab &
                                          sqlDR("pmic_vendor") & vbTab &
                                          sqlDR("pmic_partno") & vbTab &
                                          sqlDR("pmic_top_marking") & vbTab &
                                          sqlDR("pmic_org_partno") & vbTab &
                                          sqlDR("rcd_vendor") & vbTab &
                                          sqlDR("rcd_partno") & vbTab &
                                          sqlDR("rcd_top_marking") & vbTab &
                                          sqlDR("rcd_org_partno") & vbTab &
                                          sqlDR("note") & vbTab &
                                          sqlDR("tb_unique")
            GRID_Data_List.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        If GRID_Data_List.Rows.Count = 1 Then
            GRID_Data_List.AddItem("N")
            GRID_Num_Sort()
        End If

        GRID_Data_List.AutoSizeCols()
        GRID_Data_List.Redraw = True

        thread_LoadingFormEnd

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub
End Class
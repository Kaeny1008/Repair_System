Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class HistoryBLUForm
    Private Sub HistoryBLUForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        GRID_Setting()

    End Sub

    Private Sub GRID_Setting()

        With Grid_PartsChange
            .Redraw = False
            .AllowEditing = False
            .AllowFiltering = True
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.Both
            .AllowDragging = AllowDraggingEnum.None
            .AllowResizing = AllowResizingEnum.None
            .KeyActionEnter = KeyActionEnum.MoveDown
            .KeyActionTab = KeyActionEnum.MoveAcross
            .SelectionMode = SelectionModeEnum.Default
            '.Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 19
            .Cols.Fixed = 1
            .Rows.Count = 2
            .Rows.Fixed = 2
            .AllowMerging = AllowMergingEnum.FixedOnly
            .Rows(0).AllowMerging = True
            .Rows(1).AllowMerging = True
            Dim Mrng As CellRange = .GetCellRange(0, 0, 1, 0)
            Mrng.Data = "No"
            Mrng = .GetCellRange(0, 1, 1, 1)
            Mrng.Data = "CheckCOde"
            Mrng = .GetCellRange(0, 2, 1, 2)
            Mrng.Data = "자재" & vbCrLf & "구분"
            Mrng = .GetCellRange(0, 3, 1, 3)
            Mrng.Data = "작업시간"
            Mrng = .GetCellRange(0, 4, 1, 4)
            Mrng.Data = "Model"
            Mrng = .GetCellRange(0, 5, 1, 5)
            Mrng.Data = "Marking or" & vbCrLf & "Lot No."
            Mrng = .GetCellRange(0, 6, 1, 6)
            Mrng.Data = "작업Line"
            Mrng = .GetCellRange(0, 7, 1, 7)
            Mrng.Data = "작업자"
            Mrng = .GetCellRange(0, 8, 1, 8)
            Mrng.Data = "Feeder" & vbCrLf & "No"
            Mrng = .GetCellRange(0, 9, 1, 9)
            Mrng.Data = "확인" & vbCrLf & "Part No."
            Mrng = .GetCellRange(0, 10, 0, 11)
            Mrng.Data = "교환전"
            Mrng = .GetCellRange(1, 10, 1, 10)
            Mrng.Data = "Part No."
            Mrng = .GetCellRange(1, 11, 1, 11)
            Mrng.Data = "Rank"
            Mrng = .GetCellRange(0, 12, 0, 15)
            Mrng.Data = "교환후"
            Mrng = .GetCellRange(1, 12, 1, 12)
            Mrng.Data = "Part No."
            Mrng = .GetCellRange(1, 13, 1, 13)
            Mrng.Data = "Rank"
            Mrng = .GetCellRange(1, 14, 1, 14)
            Mrng.Data = "Lot No."
            Mrng = .GetCellRange(1, 15, 1, 15)
            Mrng.Data = "Q'ty"
            Mrng = .GetCellRange(0, 16, 1, 16)
            Mrng.Data = "결과"
            Mrng = .GetCellRange(0, 17, 1, 17)
            Mrng.Data = "NG 사유"
            Mrng = .GetCellRange(0, 18, 1, 18)
            Mrng.Data = "NG 확인자"
            For i = 0 To .Cols.Count - 1
                .Cols(i).AllowMerging = True
            Next i

            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter '비고란은 좌측 정렬
            .ExtendLastCol = False '마지막 열 확장
            .AutoSizeCols()
            .ShowCursor = True
            .Cols.Frozen = 0
            .FocusRect = FocusRectEnum.Light
            .Cols(1).Visible = False
            .Redraw = True
        End With

    End Sub

    Private Sub Btn_changeSearch_Click(sender As Object, e As EventArgs) Handles Btn_changeSearch.Click

        Grid_PartsChange.Redraw = False
        Grid_PartsChange.Rows.Count = 2

        DBConnect()

        Dim strSQL As String = "call USP_HISTORY(6"
        strSQL += ",'" & Format(DTP_startDate2.Value, "yyyy-MM-dd 00:00:00") & "'" 'startDate
        strSQL += ",'" & Format(DTP_endDate2.Value, "yyyy-MM-dd 23:59:59") & "'" 'endDate
        strSQL += ",null" 'customerName
        strSQL += ",'" & Tb_modelName2.Text & "'" 'modelName
        strSQL += ",null" 'factoryName
        strSQL += ",'" & Cb_workLine2.Text & "'" 'workLine
        strSQL += ",'" & TB_Marking.Text & "'" 'workSide
        strSQL += ",null" 'checkCode
        strSQL += ",'" & Cb_result.Text & "'" 'checkResult
        strSQL += ",null)" 'firstSetting

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim checkResult As String = String.Empty
            If sqlDR("CHG_RESULT") = "OK" Then
                checkResult = "정상"
            ElseIf sqlDR("CHG_RESULT") = "NG1" Then
                checkResult = "교환 전 Part No.가 다릅니다."
            ElseIf sqlDR("CHG_RESULT") = "NG2" Then
                checkResult = "교환 후 Part No.가 다릅니다."
            ElseIf sqlDR("CHG_RESULT") = "NG3" Then
                checkResult = "교환 전 Rank 오류 입니다."
            ElseIf sqlDR("CHG_RESULT") = "NG4" Then
                checkResult = "교환 후 Rank 오류 입니다."
            End If
            Dim insert As String = String.Empty
            insert = Grid_PartsChange.Rows.Count - 1 &
                vbTab & sqlDR("CHECK_CODE") &
                vbTab & sqlDR("PART_DIVISION") &
                vbTab & Format(sqlDR("CHECK_DATE"), "yyyy-MM-dd HH:mm:ss") &
                vbTab & sqlDR("MODEL_NAME") &
                vbTab & sqlDR("MARKING") &
                vbTab & sqlDR("WORK_LINE") &
                vbTab & sqlDR("WORKER") &
                vbTab & sqlDR("FEEDER_NO") &
                vbTab & sqlDR("ORG_PART_NO") &
                vbTab & sqlDR("BEF_PART_NO") &
                vbTab & sqlDR("BEF_RANK") &
                vbTab & sqlDR("CHG_PART_NO") &
                vbTab & sqlDR("CHG_RANK") &
                vbTab & sqlDR("CHG_LOT_NO") &
                vbTab & Format(sqlDR("CHG_QTY"), "#,##0") &
                vbTab & checkResult &
                vbTab & sqlDR("NG_RESULT") &
                vbTab & sqlDR("NG_CHECK_ID")
            Grid_PartsChange.AddItem(insert)
        Loop
        sqlDR.Close()

        DBClose()

        Grid_PartsChange.AutoSizeCols()
        Grid_PartsChange.Redraw = True

    End Sub

    Private Sub Grid_PartsChange_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles Grid_PartsChange.MouseDoubleClick

        If e.Button = MouseButtons.Left And Grid_PartsChange.MouseRow > 0 Then
            Grid_PartsChange.Row = Grid_PartsChange.MouseRow
            Dim checkCode As String = Grid_PartsChange(Grid_PartsChange.Row, 1)
            If Not Grid_PartsChange(Grid_PartsChange.Row, 2) = "LED" Then Exit Sub
            BLU_RankList.checkCode = checkCode
            If Not BLU_RankList.Visible Then BLU_RankList.Show()
            BLU_RankList.Focus()
        End If

    End Sub

    Private Sub HistoryBLUForm_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        Mainform.form_name = Me

    End Sub
End Class
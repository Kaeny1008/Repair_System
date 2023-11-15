Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class HistoryForm

    Dim form_msgbox_string As String = "이력보기"

    Private Sub HistoryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        GridSetting()

        'DTP_startDate.Value = Format(Now, "yyyy-MM-dd 00:00:00")

    End Sub

    Private Sub GridSetting()

        With Grid_AllPartsCheck
            .Redraw = False
            .AllowEditing = False
            .AllowFiltering = True
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.Both
            .AllowDragging = AllowDraggingEnum.None
            .AllowResizing = AllowResizingEnum.None
            .KeyActionEnter = KeyActionEnum.MoveDown
            .KeyActionTab = KeyActionEnum.MoveAcross
            .Cols.Count = 12
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            Grid_AllPartsCheck(0, 0) = "No"
            Grid_AllPartsCheck(0, 1) = "Check Code"
            Grid_AllPartsCheck(0, 2) = "작업시간"
            Grid_AllPartsCheck(0, 3) = "고객사"
            Grid_AllPartsCheck(0, 4) = "Model"
            Grid_AllPartsCheck(0, 5) = "공장명"
            Grid_AllPartsCheck(0, 6) = "작업라인"
            Grid_AllPartsCheck(0, 7) = "작업면"
            Grid_AllPartsCheck(0, 8) = "Machine No"
            Grid_AllPartsCheck(0, 9) = "작업자"
            Grid_AllPartsCheck(0, 10) = "알람횟수"
            Grid_AllPartsCheck(0, 11) = "세부이력"
            .Cols(10).ComboList = "..."
            .Cols(11).ComboList = "..."
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter '비고란은 좌측 정렬
            .ExtendLastCol = False '마지막 열 확장
            .AutoSizeCols()
            .ShowCursor = True
            .Cols.Frozen = 0
            .FocusRect = FocusRectEnum.Light
            .Redraw = True
        End With

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
            .Cols.Count = 18
            .Cols.Fixed = 1
            .Rows.Count = 2
            .Rows.Fixed = 2
            .AllowMerging = AllowMergingEnum.FixedOnly
            .Rows(0).AllowMerging = True
            .Rows(1).AllowMerging = True
            Dim Mrng As CellRange = .GetCellRange(0, 0, 1, 0)
            Mrng.Data = "No"
            Mrng = .GetCellRange(0, 1, 1, 1)
            Mrng.Data = "작업시간"
            Mrng = .GetCellRange(0, 2, 1, 2)
            Mrng.Data = "고객사"
            Mrng = .GetCellRange(0, 3, 1, 3)
            Mrng.Data = "Model"
            Mrng = .GetCellRange(0, 4, 1, 4)
            Mrng.Data = "공장명"
            Mrng = .GetCellRange(0, 5, 1, 5)
            Mrng.Data = "작업Line"
            Mrng = .GetCellRange(0, 6, 1, 6)
            Mrng.Data = "작업면"
            Mrng = .GetCellRange(0, 7, 1, 7)
            Mrng.Data = "작업자"
            Mrng = .GetCellRange(0, 8, 1, 8)
            Mrng.Data = "Machine" & vbCrLf & "No."
            Mrng = .GetCellRange(0, 9, 1, 9)
            Mrng.Data = "Feeder" & vbCrLf & "No"
            Mrng = .GetCellRange(0, 10, 1, 10)
            Mrng.Data = "확인" & vbCrLf & "Part No."
            Mrng = .GetCellRange(0, 11, 1, 11)
            Mrng.Data = "교환전" & vbCrLf & "Part No."
            Mrng = .GetCellRange(0, 12, 0, 14)
            Mrng.Data = "교환후"
            Mrng = .GetCellRange(1, 12, 1, 12)
            Mrng.Data = "Part No."
            Mrng = .GetCellRange(1, 13, 1, 13)
            Mrng.Data = "Lot No."
            Mrng = .GetCellRange(1, 14, 1, 14)
            Mrng.Data = "Q'ty"
            Mrng = .GetCellRange(0, 15, 1, 15)
            Mrng.Data = "결과"
            Mrng = .GetCellRange(0, 16, 1, 16)
            Mrng.Data = "NG 사유"
            Mrng = .GetCellRange(0, 17, 1, 17)
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
            .Redraw = True
        End With

    End Sub

    Private Sub Btn_Search_Click(sender As Object, e As EventArgs) Handles Btn_Search.Click

        Grid_AllPartsCheck.Redraw = False
        Grid_AllPartsCheck.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call USP_HISTORY(0"
        strSQL += ",'" & Format(DTP_startDate.Value, "yyyy-MM-dd 00:00:00") & "'"
        strSQL += ",'" & Format(DTP_endDate.Value, "yyyy-MM-dd 23:59:59") & "'"
        strSQL += ",'" & Tb_customerName.Text & "'"
        strSQL += ",'" & Tb_modelName.Text & "'"
        strSQL += ",'" & Cb_factoryName.Text & "'"
        strSQL += ",'" & Cb_workLine.Text & "'"
        strSQL += ",'" & Cb_workSide.Text & "'"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null)"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert As String = String.Empty
            insert = Grid_AllPartsCheck.Rows.Count &
                vbTab & sqlDR("CHECK_CODE") &
                vbTab & Format(sqlDR("CHECK_DATE"), "yyyy-MM-dd HH:mm:ss") &
                vbTab & sqlDR("CUSTOMER_NAME") &
                vbTab & sqlDR("MODEL_NAME") &
                vbTab & sqlDR("FACTORY_NAME") &
                vbTab & sqlDR("WORK_LINE") &
                vbTab & sqlDR("WORK_SIDE") &
                vbTab & sqlDR("MACHINE_NO") &
                vbTab & sqlDR("WORKER") &
                vbTab & sqlDR("NG_CHECK_COUNT")
            Grid_AllPartsCheck.AddItem(insert)
        Loop
        sqlDR.Close()

        DBClose()

        Grid_AllPartsCheck.AutoSizeCols()
        Grid_AllPartsCheck.Redraw = True

    End Sub

    Private Sub Grid_AllPartsCheck_RowColChange(sender As Object, e As EventArgs) Handles Grid_AllPartsCheck.RowColChange

        Select Case Grid_AllPartsCheck.Col
            Case 10, 11
                Grid_AllPartsCheck.AllowEditing = True
            Case Else
                Grid_AllPartsCheck.AllowEditing = False
        End Select

    End Sub

    Dim selFactoryCode As String
    Dim selFactoryCode2 As String

    Private Sub Cb_FactoryName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cb_factoryName.SelectedIndexChanged

        DBConnect()

        Dim strSQL As String = "select SUB_CODE from TB_SUB_CODE"
        strSQL += " where SUB_CODE_NAME = '" & Cb_factoryName.Text & "'"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            selFactoryCode = sqlDR("SUB_CODE")
        Loop
        sqlDR.Close()

        DBClose()

        Cb_workLine.SelectedIndex = -1

    End Sub

    Private Sub Cb_FactoryName2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cb_factoryName2.SelectedIndexChanged

        DBConnect()

        Dim strSQL As String = "select SUB_CODE from TB_SUB_CODE"
        strSQL += " where SUB_CODE_NAME = '" & Cb_factoryName2.Text & "'"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            selFactoryCode2 = sqlDR("SUB_CODE")
        Loop
        sqlDR.Close()

        DBClose()

        Cb_workLine2.SelectedIndex = -1

    End Sub

    Private Sub Cb_FactoryName_DropDown(sender As Object, e As EventArgs) Handles Cb_factoryName.DropDown,
                                                                                  Cb_factoryName2.DropDown

        Dim cb As ComboBox = CType(sender, ComboBox)

        Dim cb_old_string As String = cb.Text

        cb.Items.Clear()

        DBConnect()

        Dim strSQL As String = "select SUB_CODE_NAME from TB_SUB_CODE"
        strSQL += " where MAIN_CODE = 'MC0002' order by SUB_CODE_NAME"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            cb.Items.Add(sqlDR("SUB_CODE_NAME"))
        Loop
        sqlDR.Close()

        DBClose()

        cb.Text = cb_old_string

    End Sub

    Private Sub Cb_workLine_DropDown(sender As Object, e As EventArgs) Handles Cb_workLine.DropDown,
                                                                               Cb_workLine2.DropDown
        Dim cb As ComboBox = CType(sender, ComboBox)

        'If Cb_factoryName.Text = String.Empty Then
        '    MsgBox("공장을 먼저 선택하여 주십시오.", MsgBoxStyle.Information, form_msgbox_string)
        '    Exit Sub
        'End If

        Dim cb_old_string As String = cb.Text

        cb.Items.Clear()

        DBConnect()

        Dim strSQL As String = String.Empty

        If cb.Name = Cb_workLine.Name Then
            strSQL = "select LAST_CODE_NAME from TB_LAST_CODE"
            strSQL += " where SUB_CODE = '" & selFactoryCode & "' order by LAST_CODE_NAME"
        ElseIf cb.Name = Cb_workLine2.Name Then
            strSQL = "select LAST_CODE_NAME from TB_LAST_CODE"
            strSQL += " where SUB_CODE = '" & selFactoryCode2 & "' order by LAST_CODE_NAME"
        End If

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            cb.Items.Add(sqlDR("LAST_CODE_NAME"))
        Loop
        sqlDR.Close()

        DBClose()

        cb.Text = cb_old_string

    End Sub

    Private Sub Grid_AllPartsCheck_CellButtonClick(sender As Object, e As RowColEventArgs) Handles Grid_AllPartsCheck.CellButtonClick

        Select Case e.Col
            Case 10 '알람횟수
                AlarmHistory.checkCode = Grid_AllPartsCheck(e.Row, 1)
                If Not AlarmHistory.Visible Then AlarmHistory.Show()
                AlarmHistory.Focus()
            Case 11 '세부이력
                AllPartsCheckHistory.checkCode = Grid_AllPartsCheck(e.Row, 1)
                If Not AllPartsCheckHistory.Visible Then AllPartsCheckHistory.Show()
                AllPartsCheckHistory.Focus()
        End Select

    End Sub

    Private Sub Btn_changeSearch_Click(sender As Object, e As EventArgs) Handles Btn_changeSearch.Click

        Grid_PartsChange.Redraw = False
        Grid_PartsChange.Rows.Count = 2

        DBConnect()

        Dim firstSttingShow As String = "No"

        If CB_FirstSetting.Checked Then
            firstSttingShow = "Yes"
        End If

        Dim strSQL As String = "call USP_HISTORY(3"
        strSQL += ",'" & Format(DTP_startDate2.Value, "yyyy-MM-dd 00:00:00") & "'"
        strSQL += ",'" & Format(DTP_endDate2.Value, "yyyy-MM-dd 23:59:59") & "'"
        strSQL += ",'" & Tb_customerName2.Text & "'"
        strSQL += ",'" & Tb_modelName2.Text & "'"
        strSQL += ",'" & Cb_factoryName2.Text & "'"
        strSQL += ",'" & Cb_workLine2.Text & "'"
        strSQL += ",'" & Cb_workSide2.Text & "'"
        strSQL += ",null"
        strSQL += ",'" & Cb_result.Text & "'"
        strSQL += ",'" & firstSttingShow & "')"

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
            End If
            Dim insert As String = String.Empty
            Dim insertQty As String = String.Empty
            If IsNumeric(sqlDR("CHG_QTY")) Then
                insertQty = Format(CDbl(sqlDR("CHG_QTY")), "#,##0")
            Else
                insertQty = sqlDR("CHG_QTY")
            End If
            insert = Grid_PartsChange.Rows.Count - 1 &
                vbTab & Format(sqlDR("CHECK_DATE"), "yyyy-MM-dd HH:mm:ss") &
                vbTab & sqlDR("CUSTOMER_NAME") &
                vbTab & sqlDR("MODEL_NAME") &
                vbTab & sqlDR("FACTORY_NAME") &
                vbTab & sqlDR("WORK_LINE") &
                vbTab & sqlDR("WORK_SIDE") &
                vbTab & sqlDR("WORKER") &
                vbTab & sqlDR("MACHINE_NO") &
                vbTab & sqlDR("FEEDER_NO") &
                vbTab & sqlDR("ORG_PART_NO") &
                vbTab & sqlDR("BEF_PART_NO") &
                vbTab & sqlDR("CHG_PART_NO") &
                vbTab & sqlDR("CHG_LOT_NO") &
                vbTab & insertQty &
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

    Private Sub HistoryForm_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        Mainform.form_name = Me

    End Sub
End Class
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class BluModelManager

    'MODEL_MNG 사용 DB : MODEL_BLU_LIST
    '                    MODEL_BLU_BOM
    '                    MODEL_BLU_RANK
    '작성 시작일       : 19/02/19
    '마지막 수정일     : 20/06/15

    Dim form_Msgbox_String As String = "모델 등록/관리"

    Private Sub MODEL_MNG_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        GRID_SETTING()
        'MODEL_LIST_LOAD("", "", "")
        CB_Model_Division.SelectedIndex = 2
        SplitContainer2.Panel2Collapsed = True
        Maker_DropDown_LIST(CB_TH_Maker, "MC0003")
        Maker_DropDown_LIST(CB_PKG_Maker, "MC0003")
        APP_SHEET_DATE.CustomFormat = " "
        FLOOR_PLAN_DATE.CustomFormat = " "
        MARKING_DATE.CustomFormat = " "
        OQC_FLOOR_PLAN_DATE.CustomFormat = " "

    End Sub

    Private Sub GRID_SETTING()

        With GRID_ModelList
            .AllowEditing = True
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 17
            .Cols.Fixed = 1
            .Rows.Count = 2
            .Rows.Fixed = 1
            GRID_ModelList(0, 0) = "No"
            GRID_ModelList(0, 1) = "Model Code"
            GRID_ModelList(0, 2) = "구분"
            GRID_ModelList(0, 3) = "Model Name"
            GRID_ModelList(0, 4) = "승인원 일자"
            GRID_ModelList(0, 5) = "도면 일자"
            GRID_ModelList(0, 6) = "마킹기준 일자"
            GRID_ModelList(0, 7) = "PKG Maker"
            GRID_ModelList(0, 8) = "PKG Part No."
            GRID_ModelList(0, 9) = "PKG Q'ty"
            GRID_ModelList(0, 10) = "Intensity RANK"
            GRID_ModelList(0, 11) = "Chromaticity RANK"
            GRID_ModelList(0, 12) = "TH Maker"
            GRID_ModelList(0, 13) = "TH Part No."
            GRID_ModelList(0, 14) = "TH Q'ty"
            GRID_ModelList(0, 15) = "PCB Part No."
            GRID_ModelList(0, 16) = "비고"
            .AutoClipboard = True
            .ShowCursor = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(16).StyleNew.TextAlign = TextAlignEnum.LeftCenter '비고란은 좌측 정렬
            .AutoSizeCols()
            .Cols(2).ComboList = "개발|양산"
            '날짜 형식과 입력마스크 지정
            For i = 4 To 6
                .Cols(i).DataType = GetType(Date)
                .Cols(i).Format = "yyyy-MM-dd"
                .Cols(i).EditMask = "9999-99-99"
            Next
            '수량 입력부분 지정
            .Cols(9).DataType = GetType(Integer)
            .Cols(9).Format = "#,##0"
            .Cols(14).DataType = GetType(Integer)
            .Cols(14).Format = "#,##0"
        End With

    End Sub

    Private Sub MODEL_LIST_LOAD(ByVal division As String, ByVal model_name As String, ByVal model_run As String)

        GRID_ModelList.Redraw = False

        DBConnect()

        GRID_ModelList.Rows.Count = 1

        Dim strSql As String = String.Empty

        strSql = "call USP_MODEL_BLU_MNG('" & division & "','" & model_name & "','" & model_run & "')"

        Dim sqlCmd As New MySqlCommand(strSql, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim replace_Division As String = String.Empty
            If sqlDR("DIVISION") = "MASS" Then
                replace_Division = "양산"
            ElseIf sqlDR("DIVISION") = "DEVE" Then
                replace_Division = "개발"
            End If

            Dim row_num As String = GRID_ModelList.Rows.Count
            If sqlDR("MODEL_RUN") = "No" Then row_num = "X"

            Dim insert_String As String = row_num & vbTab &
                                          sqlDR("MODEL_CODE") & vbTab &
                                          replace_Division & vbTab &
                                          sqlDR("MODEL_NAME") & vbTab &
                                          sqlDR("APP_SHEET_DATE") & vbTab &
                                          sqlDR("FLOOR_PLAN_DATE") & vbTab &
                                          sqlDR("MARKING_DATE") & vbTab &
                                          sqlDR("PKG_MAKER") & vbTab &
                                          sqlDR("PKG_PARTNO") & vbTab &
                                          sqlDR("PKG_QTY") & vbTab &
                                          sqlDR("INTENSITY") & vbTab &
                                          sqlDR("CHROMATICITY") & vbTab &
                                          sqlDR("TH_MAKER") & vbTab &
                                          sqlDR("TH_PARTNO") & vbTab &
                                          sqlDR("TH_QTY") & vbTab &
                                          sqlDR("PCB_PARTNO") & vbTab &
                                          sqlDR("MODEL_NOTE")
            GRID_ModelList.AddItem(insert_String)

            If row_num = "X" Then
                GRID_ModelList.Rows(GRID_ModelList.Rows.Count - 1).StyleNew.ForeColor = Color.Gray
            End If
        Loop
        sqlDR.Close()

        DBClose()

        If GRID_ModelList.Rows.Count = 1 Then
            GRID_ModelList.AddItem("N")
            GRID_ModelList.Rows(GRID_ModelList.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
        End If

        GRID_ModelList.AutoSizeCols()

        GRID_ModelList.Redraw = True

    End Sub

    Private Sub MODEL_LIST_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles GRID_ModelList.KeyDown

        If e.KeyCode = 93 Then
            If GRID_ModelList.Row > 0 And GRID_ModelList.Col > 0 Then
                If GRID_ModelList(GRID_ModelList.Row, 0) = "X" Then
                    BTN_DELETE.Text = "모델사용"
                Else
                    BTN_DELETE.Text = "모델사용 안함"
                End If
                GRID_MENU.Show(GRID_ModelList, New Point(MousePosition.X - 5, MousePosition.Y - 145))
            End If
        End If

    End Sub

    Private Sub GRID_ModelList_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GRID_ModelList.MouseClick

        'If e.Button = Windows.Forms.MouseButtons.Right And MODEL_LIST.MouseRow > 0 Then
        '    MODEL_LIST.Row = MODEL_LIST.MouseRow
        '    If MODEL_LIST(MODEL_LIST.Row, 0) = "X" Then
        '        BTN_DELETE.Text = "모델사용"
        '    Else
        '        BTN_DELETE.Text = "모델사용 안함"
        '    End If
        '    GRID_MENU.Show(MODEL_LIST, New Point(e.X, e.Y))
        'End If

        If e.Button = Windows.Forms.MouseButtons.Left Then
            SplitContainer2.Panel2Collapsed = True
            BTN_Save.Enabled = False
            BTN_ModelSave.Enabled = False
        End If

    End Sub

    Private Sub MODEL_LIST_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GRID_ModelList.MouseDoubleClick

        If e.Button = Windows.Forms.MouseButtons.Left And GRID_ModelList.MouseRow > 0 Then
            CONTROL_INIT()
            BTN_Save.Enabled = True
            BTN_ModelSave.Enabled = True
            GRID_ModelList.Row = GRID_ModelList.MouseRow

            '등록(서버저장)된 모델코드이면 창이 열리도록 코딩해야함.

            DBConnect()

            Dim strSql As String = String.Empty

            strSql = "call USP_MODEL_BLU_MNG('','" & GRID_ModelList(GRID_ModelList.Row, 1) & "','')"

            Dim sqlCmd As New MySqlCommand(strSql, DBConnect1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                If sqlDR("MODEL_RUN") = "Yes" Then
                    RB_USE.Checked = True
                Else
                    RB_NOT_USE.Checked = True
                End If
                If sqlDR("DIVISION") = "MASS" Then
                    CB_ModelDivision.SelectedIndex = 1
                ElseIf sqlDR("DIVISION") = "DEVE" Then
                    CB_ModelDivision.SelectedIndex = 0
                End If
                MODEL_CODE.Text = sqlDR("MODEL_CODE")
                MODEL_NAME.Text = sqlDR("MODEL_NAME")
                If IsDBNull(sqlDR("APP_SHEET_NAME")) Then
                    APP_SHEET_FILE.Text = "등록안됨"
                    APP_SHEET_DATE.Value = Format(Now, "yyyy-MM-dd")
                    APP_SHEET_DATE.Format = DateTimePickerFormat.Custom
                Else
                    APP_SHEET_FILE.Text = sqlDR("APP_SHEET_NAME")
                    APP_SHEET_DATE.Value = Format(sqlDR("APP_SHEET_DATE"), "yyyy-MM-dd")
                    APP_SHEET_DATE.Format = DateTimePickerFormat.Short
                End If
                If IsDBNull(sqlDR("FLOOR_PLAN_NAME")) Then
                    FLOOR_PLAN_FILE.Text = "등록안됨"
                    FLOOR_PLAN_DATE.Value = Format(Now, "yyyy-MM-dd")
                    FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Custom
                Else
                    FLOOR_PLAN_FILE.Text = sqlDR("FLOOR_PLAN_NAME")
                    FLOOR_PLAN_DATE.Value = Format(sqlDR("FLOOR_PLAN_DATE"), "yyyy-MM-dd")
                    FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Short
                End If
                If IsDBNull(sqlDR("MARKING_NAME")) Then
                    MARKING_FILE.Text = "등록안됨"
                    MARKING_DATE.Value = Format(Now, "yyyy-MM-dd")
                    MARKING_DATE.Format = DateTimePickerFormat.Custom
                Else
                    MARKING_FILE.Text = sqlDR("MARKING_NAME")
                    MARKING_DATE.Value = Format(sqlDR("MARKING_DATE"), "yyyy-MM-dd")
                    MARKING_DATE.Format = DateTimePickerFormat.Short
                End If
                If IsDBNull(sqlDR("OQC_FLOOR_PLAN_NAME")) Then
                    OQC_FLOOR_PLAN_FILE.Text = "등록안됨"
                    OQC_FLOOR_PLAN_DATE.Value = Format(Now, "yyyy-MM-dd")
                    OQC_FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Custom
                Else
                    OQC_FLOOR_PLAN_FILE.Text = sqlDR("OQC_FLOOR_PLAN_NAME")
                    OQC_FLOOR_PLAN_DATE.Value = Format(sqlDR("OQC_FLOOR_PLAN_DATE"), "yyyy-MM-dd")
                    OQC_FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Short
                End If
                For i = 0 To CB_PKG_Maker.Items.Count - 1
                    If sqlDR("PKG_MAKER") = CB_PKG_Maker.Items(i).ToString Then
                        CB_PKG_Maker.SelectedIndex = i
                        Exit For
                    End If
                Next
                PKG_PARTNO.Text = sqlDR("PKG_PARTNO")
                PKG_QTY.Value = sqlDR("PKG_QTY")
                For i = 0 To CB_RankOrder.Items.Count - 1
                    If sqlDR("RANK_ORDER") = CB_RankOrder.Items(i).ToString Then
                        CB_RankOrder.SelectedIndex = i
                        Exit For
                    End If
                Next
                INTENSITY_RANK.Text = sqlDR("INTENSITY")
                CHROMATICITY_RANK.Text = sqlDR("CHROMATICITY")

                RANK_LIST_WRITE()

                For i = 0 To CB_TH_Maker.Items.Count - 1
                    If sqlDR("TH_MAKER") = CB_TH_Maker.Items(i).ToString Then
                        CB_TH_Maker.SelectedIndex = i
                        Exit For
                    End If
                Next
                TH_PARTNO.Text = sqlDR("TH_PARTNO")
                TH_QTY.Value = sqlDR("TH_QTY")
                PCB_PARTNO.Text = sqlDR("PCB_PARTNO")
                MODEL_NOTE.Text = sqlDR("MODEL_NOTE")
                PACKING_QTY.Value = sqlDR("PACKING_QTY")
                MODEL_PART_NO.Text = sqlDR("MODEL_PART_NO")
            Loop
            sqlDR.Close()

            DBClose()

            SplitContainer2.Panel2Collapsed = False
            SplitContainer2.SplitterDistance = 200
            GRID_ModelList.TopRow = GRID_ModelList.Row
            modify_model = True

            'BTN_MODEL_SAVE.Enabled = True
            app_sheet_new_upload = False
            floor_plan_new_upload = False
            marking_new_upload = False
            CB_ModelDivision.Enabled = False
        End If

    End Sub

    Private Sub RANK_LIST_WRITE()

        ListBox1.Items.Clear()

        Dim rank1() As String
        Dim rank2() As String

        If CB_RankOrder.Text = "휘도-색좌표" Then
            rank1 = Split(Replace(INTENSITY_RANK.Text, " ", ""), ",")
            rank2 = Split(Replace(CHROMATICITY_RANK.Text, " ", ""), ",")
        Else
            rank1 = Split(Replace(CHROMATICITY_RANK.Text, " ", ""), ",")
            rank2 = Split(Replace(INTENSITY_RANK.Text, " ", ""), ",")
        End If

        For i = 0 To UBound(rank1)
            Dim sel_intensity As String = rank1(i)
            For j = 0 To UBound(rank2)
                Dim sel_chromaticity As String = rank2(j)
                ListBox1.Items.Add(sel_intensity & sel_chromaticity)
            Next
        Next

    End Sub

    Dim modify_model As Boolean

    Private Function NEW_MODEL_Code(ByVal model_division As String) As String

        DBConnect()

        Dim strSql As String = "select MODEL_CODE from TB_MODEL_BLU_LIST where DIVISION = '" & model_division & "' order by MODEL_CODE limit 1"

        Dim sqlCmd As New MySqlCommand(strSql, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader
        Dim last_MODEL_Code As String = String.Empty
        NEW_MODEL_Code = String.Empty

        Do While sqlDR.Read
            last_MODEL_Code = sqlDR("MODEL_CODE")
        Loop
        sqlDR.Close()

        DBClose()

        If last_MODEL_Code = String.Empty Then
            If model_division = "MASS" Then
                NEW_MODEL_Code = "MM0001"
            ElseIf model_division = "DEVE" Then
                NEW_MODEL_Code = "MD0001"
            End If
        Else
            NEW_MODEL_Code = last_MODEL_Code.Substring(0, 2) & Format(CInt(CInt(last_MODEL_Code.Substring(2, 4)) + 1), "0000")
        End If

        For i = 1 To GRID_ModelList.Rows.Count - 1
            If NEW_MODEL_Code = GRID_ModelList(i, 1) Then
                NEW_MODEL_Code = NEW_MODEL_Code.Substring(0, 2) & Format(CInt(CInt(NEW_MODEL_Code.Substring(2, 4)) + 1), "0000")
            End If
        Next

        Return NEW_MODEL_Code

    End Function

    Private Sub MODEL_LIST_AfterEdit(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles GRID_ModelList.AfterEdit

        If IsNumeric(GRID_ModelList(e.Row, 0)) Then
            GRID_ModelList(e.Row, 0) = "M"
            GRID_ModelList.Rows(e.Row).StyleNew.ForeColor = Color.Red
        End If

        Select Case e.Col
            Case 2
                '신규 모델이라면
                If GRID_ModelList(e.Row, 0) = "N" Then
                    '모델코드 작성하는 부분 만들어야함
                    If GRID_ModelList(e.Row, e.Col) = "개발" Then
                        GRID_ModelList(e.Row, 1) = NEW_MODEL_Code("DEVE")
                    ElseIf GRID_ModelList(e.Row, e.Col) = "양산" Then
                        GRID_ModelList(e.Row, 1) = NEW_MODEL_Code("MASS")
                    End If
                Else
                    GRID_ModelList(e.Row, e.Col) = before_grid_text
                    MsgBox("모델 구분을 변경할 수 없습니다.", MsgBoxStyle.Information, form_Msgbox_String)
                End If
        End Select

        GRID_ModelList.AutoSizeCols()

    End Sub

    Dim before_grid_text As String

    Private Sub MODEL_LIST_BeforeEdit(ByVal sender As Object, ByVal e As C1.Win.C1FlexGrid.RowColEventArgs) Handles GRID_ModelList.BeforeEdit

        before_grid_text = GRID_ModelList(e.Row, e.Col)

    End Sub

    Private Sub BTN_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_ADD.Click

        GRID_ModelList.AddItem("N")
        GRID_ModelList.TopRow = GRID_ModelList.Rows.Count - 1
        GRID_ModelList.Rows(GRID_ModelList.Rows.Count - 1).StyleNew.ForeColor = Color.Blue

    End Sub

    Private Sub BTN_DELETE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_DELETE.Click

        If Not GRID_ModelList(GRID_ModelList.Row, 0) = "X" Then
            GRID_ModelList(GRID_ModelList.Row, 0) = "X"
            GRID_ModelList.Rows(GRID_ModelList.Rows.Count - 1).StyleNew.ForeColor = Color.Gray
        Else
            GRID_ModelList(GRID_ModelList.Row, 0) = "O"
            GRID_ModelList.Rows(GRID_ModelList.Rows.Count - 1).StyleNew.ForeColor = Color.Blue
        End If

    End Sub

    Private Sub MODEL_LIST_RowColChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles GRID_ModelList.RowColChange

        'If Not MainForm.LOGIN_ID.Text.ToUpper = "Login ID : ADMIN".ToUpper Then
        '    If User_Authority_Write_Check("3") = False Then
        '        MODEL_LIST.AllowEditing = False
        '        Exit Sub
        '    End If
        'End If

        Select Case GRID_ModelList.Col
            Case 1
                GRID_ModelList.AllowEditing = False
            Case Else
                GRID_ModelList.AllowEditing = True
        End Select

        Select Case GRID_ModelList.Col
            Case 7, 12 'PKG, TH Maker
                COMBO_LIST_LOAD("MC0003", GRID_ModelList.Col)
        End Select

        If GRID_ModelList.Row < 1 Then Exit Sub
        '신규가 아니라면 수정금지
        If Not GRID_ModelList(GRID_ModelList.Row, 0) = "N" Then
            GRID_ModelList.AllowEditing = False
        Else
            GRID_ModelList.AllowEditing = True
        End If

    End Sub

    Private Sub COMBO_LIST_LOAD(ByVal division_code As String, ByVal col_no As Integer)

        DBConnect()

        Dim strSql As String = "select * from TB_SUB_CODE where MAIN_CODE = '" & division_code & "'"

        Dim sqlCmd As New MySqlCommand(strSql, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader
        Dim combo_list As String = String.Empty

        Do While sqlDR.Read
            If combo_list = String.Empty Then
                combo_list = sqlDR("SUB_CODE_NAME")
            Else
                combo_list += "|" & sqlDR("SUB_CODE_NAME")
            End If
        Loop
        sqlDR.Close()

        DBClose()

        GRID_ModelList.Cols(col_no).ComboList = combo_list

    End Sub

    Private Sub BTN_SAVE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Save.Click

        If MsgBox("저장 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = DBConnect1.BeginTransaction

        Try
            For i = 1 To GRID_ModelList.Rows.Count - 1
                Dim model_Code As String = GRID_ModelList(i, 1)
                Dim model_Division As String = GRID_ModelList(i, 2)
                If model_Division = "개발" Then
                    model_Division = "DEVE"
                ElseIf model_Division = "양산" Then
                    model_Division = "MASS"
                End If
                Dim model_Name As String = GRID_ModelList(i, 3)
                Dim model_APP_Sheet_Date As String = GRID_ModelList(i, 4)
                Dim model_Floor_Plan_Date As String = GRID_ModelList(i, 5)
                Dim model_Marking_Date As String = GRID_ModelList(i, 6)
                Dim model_PKG_Maker As String = GRID_ModelList(i, 7)
                Dim model_PKG_PartNo As String = GRID_ModelList(i, 8)
                Dim model_PKG_QTY As String = Replace(GRID_ModelList(i, 9), ",", "")
                Dim model_Intensity_RANK As String = GRID_ModelList(i, 10)
                Dim model_Chromaticity_RANK As String = GRID_ModelList(i, 11)
                Dim model_TH_Maker As String = GRID_ModelList(i, 12)
                Dim model_TH_PartNo As String = GRID_ModelList(i, 13)
                Dim model_TH_QTY As String = Replace(GRID_ModelList(i, 14), ",", "")
                Dim model_PCB_PartNo As String = GRID_ModelList(i, 15)
                Dim model_Note As String = GRID_ModelList(i, 16)
                If model_PKG_QTY = String.Empty Then model_PKG_QTY = 0
                If model_TH_QTY = String.Empty Then model_TH_QTY = 0
                If GRID_ModelList(i, 0) = "N" Then '신규등록
                    strSQL += "insert into TB_MODEL_BLU_LIST values("
                    strSQL += "'Yes'"
                    strSQL += ",'" & model_Division & "'"
                    strSQL += ",'" & model_Code & "'"
                    strSQL += ",'" & model_Name & "'"
                    strSQL += ",'" & model_APP_Sheet_Date & "'"
                    strSQL += ",'" & model_Floor_Plan_Date & "'"
                    strSQL += ",'" & model_Marking_Date & "'"
                    strSQL += ",'" & model_Note & "');"
                    strSQL += "insert into TB_MODEL_BLU_BOM values("
                    strSQL += "'" & model_Code & "'"
                    strSQL += ",'" & model_PKG_Maker & "'"
                    strSQL += ",'" & model_PKG_PartNo & "'"
                    strSQL += ",'" & model_PKG_QTY & "'"
                    strSQL += ",'" & model_TH_Maker & "'"
                    strSQL += ",'" & model_TH_PartNo & "'"
                    strSQL += ",'" & model_TH_QTY & "'"
                    strSQL += ",'" & model_PCB_PartNo & "');"
                    strSQL += "insert into TB_MODEL_BLU_RANK values("
                    strSQL += "'" & model_Code & "'"
                    strSQL += ",'" & model_Intensity_RANK & "'"
                    strSQL += ",'" & model_Chromaticity_RANK & "');"
                ElseIf GRID_ModelList(i, 0) = "X" Then '삭제
                    strSQL += "update TB_MODEL_BLU_LIST set MODEL_RUN = 'No' where MODEL_CODE = '" & model_Code & "';"
                ElseIf GRID_ModelList(i, 0) = "O" Then '삭제
                    strSQL += "update TB_MODEL_BLU_LIST set MODEL_RUN = 'Yes' where MODEL_CODE = '" & model_Code & "';"
                ElseIf GRID_ModelList(i, 0) = "M" Then '수정 <------ 수정부문은 추가할지 결정해야함.

                End If
            Next

            sqlCmd = New MySqlCommand(strSQL, DBConnect1)
            sqlCmd.Transaction = sqlTran
            sqlCmd.ExecuteNonQuery()

            sqlTran.Commit()
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        MsgBox("저장완료.", MsgBoxStyle.Information, form_Msgbox_String)

        Dim division_text As String = String.Empty
        Dim model_run As String = String.Empty

        If CB_Model_Division.Text = "개발" Then
            division_text = "DEVE"
        ElseIf CB_Model_Division.Text = "양산" Then
            division_text = "MASS"
        ElseIf CB_Model_Division.Text = "모두" Then
            division_text = ""
        End If

        If RB_Model_ALL.Checked Then
            model_run = ""
        ElseIf RB_Model_Not_Use.Checked Then
            model_run = "No"
        ElseIf RB_Model_Use.Checked Then
            model_run = "Yes"
        End If

        MODEL_LIST_LOAD(division_text, TB_Model_Name.Text, model_run)

    End Sub

    Private Sub BTN_SEARCH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Search.Click

        BTN_Save.Enabled = False
        BTN_ModelSave.Enabled = False
        SplitContainer2.Panel2Collapsed = True
        SplitContainer2.Panel1Collapsed = False
        SplitContainer1.Panel1Collapsed = False

        Dim division_text As String = String.Empty
        Dim model_run As String = String.Empty

        If CB_Model_Division.Text = "개발" Then
            division_text = "DEVE"
        ElseIf CB_Model_Division.Text = "양산" Then
            division_text = "MASS"
        ElseIf CB_Model_Division.Text = "모두" Then
            division_text = ""
        End If

        If RB_Model_ALL.Checked Then
            model_run = ""
        ElseIf RB_Model_Not_Use.Checked Then
            model_run = "No"
        ElseIf RB_Model_Use.Checked Then
            model_run = "Yes"
        End If

        MODEL_LIST_LOAD(division_text, TB_Model_Name.Text, model_run)

    End Sub

    Private Sub TB_Model_Name_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TB_Model_Name.GotFocus

        TB_Model_Name.SelectAll()

    End Sub

    Private Sub TB_Model_Name_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TB_Model_Name.KeyDown

        If e.KeyCode = 13 And Not TB_Model_Name.Text = String.Empty Then
            BTN_SEARCH_Click(Nothing, Nothing)
        End If

    End Sub

    Private Sub TB_Model_Name_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TB_Model_Name.MouseClick

        TB_Model_Name.SelectAll()

    End Sub

    Private Sub MODEL_LIST_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GRID_ModelList.Click

    End Sub

    Private Sub MODEL_FILE_FOLDER(ByVal model_code As String,
                                  ByVal folder_division As String,
                                  ByVal file_date As String)

        'ROOT 폴더에서 MODEL 폴더가 존재하는지 검색 후 존재하지 않으면 생성
        Dim ftpFolder As String = ftpFolderSearch(ftpUrl & "/")

        Dim DirectoryName() As String = Split(ftpFolder, vbCrLf)
        Dim DirectoryExists As Boolean

        For i = 0 To UBound(DirectoryName) - 1
            If "MODEL" = Trim(DirectoryName(i)) Then
                DirectoryExists = True
                Exit For
            Else
                DirectoryExists = False
            End If
        Next

        If DirectoryExists = False Then
            ftpFolderMake(ftpUrl & "/MODEL")
        End If

        'ROOT/MODEL 폴더에서 해당 폴더를 검색 후 존재하지 않으면 생성
        DirectoryExists = False '초기화

        ftpFolder = ftpFolderSearch(ftpUrl & "/MODEL/")

        DirectoryName = Split(ftpFolder, vbCrLf)

        For i = 0 To UBound(DirectoryName) - 1
            If model_code = Trim(DirectoryName(i)) Then
                DirectoryExists = True
                Exit For
            Else
                DirectoryExists = False
            End If
        Next

        If DirectoryExists = False Then
            ftpFolderMake(ftpUrl & "/MODEL/" & model_code)
        End If

        'ROOT/MODEL/folder_division 폴더에서 해당 폴더를 검색 후 존재하지 않으면 생성
        DirectoryExists = False '초기화

        ftpFolder = ftpFolderSearch(ftpUrl & "/MODEL/" & model_code & "/")

        DirectoryName = Split(ftpFolder, vbCrLf)

        For i = 0 To UBound(DirectoryName) - 1
            If folder_division = Trim(DirectoryName(i)) Then
                DirectoryExists = True
                Exit For
            Else
                DirectoryExists = False
            End If
        Next

        If DirectoryExists = False Then
            ftpFolderMake(ftpUrl & "/MODEL/" & model_code & "/" & folder_division)
        End If

        'ROOT/MODEL/folder_division/file_date 폴더에서 해당 폴더를 검색 후 존재하지 않으면 생성
        DirectoryExists = False '초기화
        ftpFolder = ftpFolderSearch(ftpUrl & "/MODEL/" & model_code & "/" & folder_division & "/")

        DirectoryName = Split(ftpFolder, vbCrLf)

        For i = 0 To UBound(DirectoryName) - 1
            If file_date = Trim(DirectoryName(i)) Then
                DirectoryExists = True
                Exit For
            Else
                DirectoryExists = False
            End If
        Next

        If DirectoryExists = False Then
            ftpFolderMake(ftpUrl & "/MODEL/" & model_code & "/" & folder_division & "/" & file_date)
        End If
        ''폴더(생성)
        'MODEL_FILE_FOLDER(model_code.Text, "APP_SHEET", Format(APP_SHEET_DATE.Value, "yyyyMMdd"))

        ''파일(업로드)
        'ftpFileUpload(ftpUrl & "/MODEL/" & model_code.Text & "/APP_SHEET/" & Format(APP_SHEET_DATE.Value, "yyyyMMdd"), _
        '                  TEMP_FILE_Folder & "\" & APP_SHEET_FILE.Text)

    End Sub

    Private Sub FILE_SELECT(ByVal file_division As String)

        Dim openFile As New System.Windows.Forms.OpenFileDialog

        With openFile
            .Filter = "ALL Files (*.*)|*.*"
            .AddExtension() = True
        End With

        openFile.ShowDialog()

        If openFile.FileName = "" Then Exit Sub

        Dim TEMP_FILE_Folder As String = Application.StartupPath & "\TEMP_FILE"

        '좌표폴더가 없으면 만들기
        If My.Computer.FileSystem.DirectoryExists(TEMP_FILE_Folder) = False Then
            My.Computer.FileSystem.CreateDirectory(TEMP_FILE_Folder)
        End If

        Dim control_name As TextBox = Nothing

        If file_division = "APP_SHEET" Then
            control_name = APP_SHEET_FILE
        ElseIf file_division = "FLOOR_PLAN" Then
            control_name = FLOOR_PLAN_FILE
        ElseIf file_division = "MARKING" Then
            control_name = MARKING_FILE
        ElseIf file_division = "OQC_FLOOR_PLAN" Then
            control_name = OQC_FLOOR_PLAN_FILE
        End If

        '기존선택한 좌표파일이 있다면..
        'If Not control_name.Text = "" Then
        If My.Computer.FileSystem.FileExists(TEMP_FILE_Folder & "\" & control_name.Text) Then '기존좌표파일이 좌표폴더에 있다면 삭제
            My.Computer.FileSystem.DeleteFile(TEMP_FILE_Folder & "\" & control_name.Text)
        End If
        'End If

        control_name.Text = openFile.FileName '새파일로 대체 
        '임시폴더 복사
        Dim BOMFIle() As String = Split(control_name.Text, "\")
        My.Computer.FileSystem.CopyFile(control_name.Text, TEMP_FILE_Folder & "\" & BOMFIle(UBound(BOMFIle)), True)
        control_name.Text = BOMFIle(UBound(BOMFIle))

    End Sub

    Dim app_sheet_new_upload As Boolean
    Dim floor_plan_new_upload As Boolean
    Dim marking_new_upload As Boolean
    Dim oqc_floor_plan_new_upload As Boolean

    Private Sub APP_SHEET_UPLOAD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles APP_SHEET_UPLOAD.Click

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "승인원 업로드")
            Exit Sub
        End If

        If APP_SHEET_DATE.Format = DateTimePickerFormat.Custom Then
            MsgBox("승인원 일자를 선택하지 않았습니다.", MsgBoxStyle.Information, "승인원 업로드")
            Exit Sub
        End If

        If MsgBox(Format(APP_SHEET_DATE.Value, "yyyy-MM-dd") & " 일자의 승인원 파일을 첨부하시겠습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "승인원 파일 첨부") = MsgBoxResult.No Then Exit Sub

        Dim TEMP_FILE_Folder As String = Application.StartupPath & "\TEMP_FILE"

        '첨부파일 선택
        FILE_SELECT("APP_SHEET")

        '폴더 생성
        'MODEL_FILE_FOLDER(MODEL_CODE.Text, "APP_SHEET", Format(DateTimePicker1.Value, "yyyyMMdd"))

        '파일 업로드
        'ftpFileUpload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/APP_SHEET/" & Format(DateTimePicker1.Value, "yyyyMMdd"), _
        '                  TEMP_FILE_Folder & "\" & APP_SHEET_FILE.Text)

        '임시저장 파일 삭제
        'My.Computer.FileSystem.DeleteFile(TEMP_FILE_Folder & "\" & APP_SHEET_FILE.Text)

        'MsgBox("File Select Success", MsgBoxStyle.Information, "승인원 업로드")

        app_sheet_new_upload = True

    End Sub

    Private Sub APP_SHEET_DOWNLOAD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles APP_SHEET_DOWNLOAD.Click

        If APP_SHEET_FILE.Text = "등록안됨" Then
            MsgBox("승인원이 등록되지 않았습니다.", MsgBoxStyle.Information, "승인원 다운로드")
            Exit Sub
        End If

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "승인원 다운로드")
            Exit Sub
        End If

        Dim saveFile As New System.Windows.Forms.FolderBrowserDialog

        saveFile.ShowDialog()

        If saveFile.SelectedPath = "" Then Exit Sub

        ftpFileDownload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/APP_SHEET/" & Format(APP_SHEET_DATE.Value, "yyyyMMdd"), saveFile.SelectedPath, APP_SHEET_FILE.Text)

        MsgBox("File Download Success", MsgBoxStyle.Information, "승인원 다운로드")

    End Sub

    Private Sub APP_SHEET_DELETE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles APP_SHEET_DELETE.Click

        If APP_SHEET_FILE.Text = "등록안됨" Then
            MsgBox("승인원이 등록되지 않았습니다.", MsgBoxStyle.Information, "승인원 삭제")
            Exit Sub
        End If

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "승인원 삭제")
            Exit Sub
        End If

        If MsgBox("첨부된 승인원 파일을 삭제 하시겠습니까?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "승인원 파일 삭제") = MsgBoxResult.No Then Exit Sub

        '파일유지
        'ftpFileDelete(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/APP_SHEET/" & Format(APP_SHEET_DATE.Value, "yyyyMMdd"), APP_SHEET_FILE.Text)

        APP_SHEET_FILE.Text = "등록안됨"

        APP_SHEET_DATE.Value = Format(Now, "yyyy-MM-dd")
        APP_SHEET_DATE.Format = DateTimePickerFormat.Custom

        MsgBox("File Delete Success" & vbCrLf & "모델 저장시 적용됩니다.", MsgBoxStyle.Information, "승인원 삭제")

    End Sub

    Private Sub FLOOR_PLAN_UPLOAD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FLOOR_PLAN_UPLOAD.Click

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "도면 업로드")
            Exit Sub
        End If

        If FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Custom Then
            MsgBox("도면 일자를 선택하지 않았습니다.", MsgBoxStyle.Information, "도면 업로드")
            Exit Sub
        End If

        If MsgBox(Format(FLOOR_PLAN_DATE.Value, "yyyy-MM-dd") & " 일자의 도면 파일을 첨부하시겠습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "도면 파일 첨부") = MsgBoxResult.No Then Exit Sub

        Dim TEMP_FILE_Folder As String = Application.StartupPath & "\TEMP_FILE"

        '첨부파일 선택
        FILE_SELECT("FLOOR_PLAN")

        '폴더 생성
        'MODEL_FILE_FOLDER(MODEL_CODE.Text, "FLOOR_PLAN", Format(DateTimePicker1.Value, "yyyyMMdd"))

        '파일 업로드
        'ftpFileUpload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/FLOOR_PLAN/" & Format(DateTimePicker1.Value, "yyyyMMdd"), _
        '                  TEMP_FILE_Folder & "\" & FLOOR_PLAN_FILE.Text)

        '임시저장 파일 삭제
        'My.Computer.FileSystem.DeleteFile(TEMP_FILE_Folder & "\" & FLOOR_PLAN_FILE.Text)

        'MsgBox("File Select Success", MsgBoxStyle.Information, "도면 업로드")

        floor_plan_new_upload = True

    End Sub

    Private Sub FLOOR_PLAN_DOWNLOAD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FLOOR_PLAN_DOWNLOAD.Click

        If FLOOR_PLAN_FILE.Text = "등록안됨" Then
            MsgBox("도면이 등록되지 않았습니다.", MsgBoxStyle.Information, "도면 다운로드")
            Exit Sub
        End If

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "도면 다운로드")
            Exit Sub
        End If

        Dim saveFile As New System.Windows.Forms.FolderBrowserDialog

        saveFile.ShowDialog()

        If saveFile.SelectedPath = "" Then Exit Sub

        ftpFileDownload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/FLOOR_PLAN/" & Format(FLOOR_PLAN_DATE.Value, "yyyyMMdd"), saveFile.SelectedPath, FLOOR_PLAN_FILE.Text)

        MsgBox("File Download Success", MsgBoxStyle.Information, "도면 다운로드")

    End Sub

    Private Sub FLOOR_PLAN_DELETE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FLOOR_PLAN_DELETE.Click

        If FLOOR_PLAN_FILE.Text = "등록안됨" Then
            MsgBox("도면이 등록되지 않았습니다.", MsgBoxStyle.Information, "도면 삭제")
            Exit Sub
        End If

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "도면 삭제")
            Exit Sub
        End If

        If MsgBox("첨부된 도면 파일을 삭제 하시겠습니까?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "도면 파일 삭제") = MsgBoxResult.No Then Exit Sub

        '파일유지
        'ftpFileDelete(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/FLOOR_PLAN/" & Format(APP_SHEET_DATE.Value, "yyyyMMdd"), FLOOR_PLAN_FILE.Text)

        FLOOR_PLAN_FILE.Text = "등록안됨"

        FLOOR_PLAN_DATE.Value = Format(Now, "yyyy-MM-dd")
        FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Custom

        MsgBox("File Delete Success" & vbCrLf & "모델 저장시 적용됩니다.", MsgBoxStyle.Information, "도면 삭제")

    End Sub

    Private Sub MARKING_UPLOAD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MARKING_UPLOAD.Click

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "마킹기준 업로드")
            Exit Sub
        End If

        If MARKING_DATE.Format = DateTimePickerFormat.Custom Then
            MsgBox("마킹기준 일자를 선택하지 않았습니다.", MsgBoxStyle.Information, "마킹기준 업로드")
            Exit Sub
        End If

        If MsgBox(Format(MARKING_DATE.Value, "yyyy-MM-dd") & " 일자의 마킹기준 파일을 첨부하시겠습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "마킹기준 파일 첨부") = MsgBoxResult.No Then Exit Sub

        Dim TEMP_FILE_Folder As String = Application.StartupPath & "\TEMP_FILE"

        '첨부파일 선택
        FILE_SELECT("MARKING")

        '폴더 생성
        'MODEL_FILE_FOLDER(MODEL_CODE.Text, "MARKING", Format(DateTimePicker1.Value, "yyyyMMdd"))

        '파일 업로드
        'ftpFileUpload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/MARKING/" & Format(DateTimePicker1.Value, "yyyyMMdd"), _
        '                  TEMP_FILE_Folder & "\" & MARKING_FILE.Text)

        '임시저장 파일 삭제
        'My.Computer.FileSystem.DeleteFile(TEMP_FILE_Folder & "\" & MARKING_FILE.Text)

        'MsgBox("File Select Success", MsgBoxStyle.Information, "마킹기준 업로드")

        marking_new_upload = True

    End Sub

    Private Sub MARKING_DOWNLOAD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MARKING_DOWNLOAD.Click

        If MARKING_FILE.Text = "등록안됨" Then
            MsgBox("마킹기준이 등록되지 않았습니다.", MsgBoxStyle.Information, "마킹기준 다운로드")
            Exit Sub
        End If

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "마킹기준 다운로드")
            Exit Sub
        End If

        Dim saveFile As New System.Windows.Forms.FolderBrowserDialog

        saveFile.ShowDialog()

        If saveFile.SelectedPath = "" Then Exit Sub

        ftpFileDownload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/MARKING/" & Format(MARKING_DATE.Value, "yyyyMMdd"), saveFile.SelectedPath, MARKING_FILE.Text)

        MsgBox("File Download Success", MsgBoxStyle.Information, "마킹기준 다운로드")

    End Sub

    Private Sub MARKING_DELETE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MARKING_DELETE.Click

        If MARKING_FILE.Text = "등록안됨" Then
            MsgBox("마킹기준이 등록되지 않았습니다.", MsgBoxStyle.Information, "마킹기준 삭제")
            Exit Sub
        End If

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "마킹기준 삭제")
            Exit Sub
        End If

        If MsgBox("첨부된 마킹기준 파일을 삭제 하시겠습니까?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "마킹 파일 삭제") = MsgBoxResult.No Then Exit Sub

        '파일유지
        'ftpFileDelete(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/MARKING/" & Format(APP_SHEET_DATE.Value, "yyyyMMdd"), MARKING_FILE.Text)

        MARKING_FILE.Text = "등록안됨"

        MARKING_DATE.Value = Format(Now, "yyyy-MM-dd")
        MARKING_DATE.Format = DateTimePickerFormat.Custom

        MsgBox("File Delete Success" & vbCrLf & "모델 저장시 적용됩니다.", MsgBoxStyle.Information, "마킹기준 삭제")

    End Sub

    Private Sub OFP_UPLOAD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OFP_UPLOAD.Click

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "도해 업로드")
            Exit Sub
        End If

        If OQC_FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Custom Then
            MsgBox("도해 일자를 선택하지 않았습니다.", MsgBoxStyle.Information, "도해 업로드")
            Exit Sub
        End If

        If MsgBox(Format(OQC_FLOOR_PLAN_DATE.Value, "yyyy-MM-dd") & " 일자의 도해 파일을 첨부하시겠습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "도해 파일 첨부") = MsgBoxResult.No Then Exit Sub

        Dim TEMP_FILE_Folder As String = Application.StartupPath & "\TEMP_FILE"

        '첨부파일 선택
        FILE_SELECT("OQC_FLOOR_PLAN")

        '폴더 생성
        'MODEL_FILE_FOLDER(MODEL_CODE.Text, "OFP", Format(DateTimePicker1.Value, "yyyyMMdd"))

        '파일 업로드
        'ftpFileUpload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/OFP/" & Format(DateTimePicker1.Value, "yyyyMMdd"), _
        '                  TEMP_FILE_Folder & "\" & OFP_FILE.Text)

        '임시저장 파일 삭제
        'My.Computer.FileSystem.DeleteFile(TEMP_FILE_Folder & "\" & OFP_FILE.Text)

        'MsgBox("File Select Success", MsgBoxStyle.Information, "마킹기준 업로드")

        oqc_floor_plan_new_upload = True

    End Sub

    Private Sub OFP_DOWNLOAD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OFP_DOWNLOAD.Click

        If OQC_FLOOR_PLAN_FILE.Text = "등록안됨" Then
            MsgBox("도해가 등록되지 않았습니다.", MsgBoxStyle.Information, "도해 다운로드")
            Exit Sub
        End If

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "도해 다운로드")
            Exit Sub
        End If

        Dim saveFile As New System.Windows.Forms.FolderBrowserDialog

        saveFile.ShowDialog()

        If saveFile.SelectedPath = "" Then Exit Sub

        ftpFileDownload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/OQC_FLOOR_PLAN/" & Format(OQC_FLOOR_PLAN_DATE.Value, "yyyyMMdd"), saveFile.SelectedPath, OQC_FLOOR_PLAN_FILE.Text)

        MsgBox("File Download Success", MsgBoxStyle.Information, "도해 다운로드")

    End Sub

    Private Sub OFP_DELETE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OFP_DELETE.Click

        If OQC_FLOOR_PLAN_FILE.Text = "등록안됨" Then
            MsgBox("도해가 등록되지 않았습니다.", MsgBoxStyle.Information, "도해 삭제")
            Exit Sub
        End If

        If MODEL_CODE.Text = String.Empty Then
            MsgBox("모델 코드가 생성되지 않았습니다.", MsgBoxStyle.Information, "도해 삭제")
            Exit Sub
        End If

        If MsgBox("첨부된 도해 파일을 삭제 하시겠습니까?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "도해 파일 삭제") = MsgBoxResult.No Then Exit Sub

        '파일유지
        'ftpFileDelete(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/OFP/" & Format(APP_SHEET_DATE.Value, "yyyyMMdd"), OFP_FILE.Text)

        OQC_FLOOR_PLAN_FILE.Text = "등록안됨"

        OQC_FLOOR_PLAN_DATE.Value = Format(Now, "yyyy-MM-dd")
        OQC_FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Custom

        MsgBox("File Delete Success" & vbCrLf & "모델 저장시 적용됩니다.", MsgBoxStyle.Information, "도해 삭제")

    End Sub

    Private Sub MODEL_DIVISION_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_ModelDivision.SelectedIndexChanged

        If modify_model = False Then

            If CB_ModelDivision.Text = "개발" Then
                MODEL_CODE.Text = NEW_MODEL_Code("DEVE")
            ElseIf CB_ModelDivision.Text = "양산" Then
                MODEL_CODE.Text = NEW_MODEL_Code("MASS")
            End If

            MODEL_NAME.SelectAll()
            MODEL_NAME.Focus()

        End If

    End Sub

    Private Sub CONTROL_INIT()

        RB_USE.Checked = True
        CB_ModelDivision.Items.Clear()
        CB_ModelDivision.Items.Add("개발")
        CB_ModelDivision.Items.Add("양산")
        MODEL_CODE.Text = String.Empty
        MODEL_NAME.Text = String.Empty
        MODEL_PART_NO.Text = String.Empty
        APP_SHEET_DATE.Value = Format(Now, "yyyy-MM-dd")
        APP_SHEET_DATE.Format = DateTimePickerFormat.Custom
        APP_SHEET_FILE.Text = String.Empty
        FLOOR_PLAN_DATE.Value = Format(Now, "yyyy-MM-dd")
        FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Custom
        FLOOR_PLAN_FILE.Text = String.Empty
        MARKING_DATE.Value = Format(Now, "yyyy-MM-dd")
        MARKING_DATE.Format = DateTimePickerFormat.Custom
        MARKING_FILE.Text = String.Empty
        OQC_FLOOR_PLAN_DATE.Value = Format(Now, "yyyy-MM-dd")
        OQC_FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Custom
        OQC_FLOOR_PLAN_FILE.Text = String.Empty
        CB_PKG_Maker_DropDown(Nothing, Nothing)
        PKG_PARTNO.Text = String.Empty
        PKG_QTY.Value = 0
        CB_TH_Maker_DropDown(Nothing, Nothing)
        TH_PARTNO.Text = String.Empty
        TH_QTY.Value = 0
        PCB_PARTNO.Text = String.Empty
        CB_RankOrder.Items.Clear()
        CB_RankOrder.Items.Add("휘도-색좌표")
        CB_RankOrder.Items.Add("색좌표-휘도")
        INTENSITY_RANK.Text = String.Empty
        CHROMATICITY_RANK.Text = String.Empty
        MODEL_NOTE.Text = String.Empty
        ListBox1.Items.Clear()

    End Sub

    Private Sub BTN_NEW_MODEL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_NewModel.Click

        CONTROL_INIT()
        SplitContainer2.Panel2Collapsed = False
        SplitContainer2.Panel1Collapsed = True
        SplitContainer1.Panel1Collapsed = True
        'SplitContainer2.SplitterDistance = 0
        BTN_NewModel.Enabled = False
        BTN_NewCancel.Enabled = True
        BTN_ModelSave.Enabled = True
        modify_model = False
        CB_ModelDivision.Enabled = True

    End Sub

    Private Sub BTN_NEW_CANCEL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_NewCancel.Click

        SplitContainer2.Panel2Collapsed = True
        SplitContainer2.Panel1Collapsed = False
        SplitContainer1.Panel1Collapsed = False
        BTN_NewModel.Enabled = True
        BTN_NewCancel.Enabled = False
        BTN_ModelSave.Enabled = False

    End Sub

    Private Sub BTN_MODEL_SAVE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_ModelSave.Click

        If CB_ModelDivision.Text = String.Empty Then
            MsgBox("모델 구분이 선택되지 않았습니다.", MsgBoxStyle.Information, "모델 구분 선택")
            Exit Sub
        End If

        If MODEL_NAME.Text = String.Empty Then
            MsgBox("모델명이 입력되지 않았습니다.", MsgBoxStyle.Information, "모델명 입력")
            Exit Sub
        End If

        If CB_PKG_Maker.Text = String.Empty Then
            MsgBox("PKG 제조사가 입력되지 않았습니다.", MsgBoxStyle.Information, "PKG 제조사 선택")
            Exit Sub
        End If

        If PKG_PARTNO.Text = String.Empty Then
            MsgBox("PKG Part No.가 입력되지 않았습니다.", MsgBoxStyle.Information, "PKG Part No. 입력")
            Exit Sub
        End If

        'If PKG_QTY.Value = 0 Then
        '    MsgBox("PKG 사용수량이 입력되지 않았습니다.", MsgBoxStyle.Information, "PKG 사용수량 입력")
        '    Exit Sub
        'End If

        If PCB_PARTNO.Text = String.Empty Then
            MsgBox("PCB Part No.가 입력되지 않았습니다.", MsgBoxStyle.Information, "PCB Part No. 입력")
            Exit Sub
        End If

        If CB_RankOrder.Text = String.Empty Then
            MsgBox("RANK 순서가 선택되지 않았습니다.", MsgBoxStyle.Information, "RANK 순서 선택")
            Exit Sub
        End If

        If INTENSITY_RANK.Text = String.Empty Then
            MsgBox("Intensity Rank 가 입력되지 않았습니다.", MsgBoxStyle.Information, "Intensity Rank 입력")
            Exit Sub
        End If

        If CHROMATICITY_RANK.Text = String.Empty Then
            MsgBox("Chromaticity Rank가 입력되지 않았습니다.", MsgBoxStyle.Information, "Chromaticity Rank 입력")
            Exit Sub
        End If

        'If PACKING_QTY.Value = 0 Then
        '    MsgBox("최대 Packing 수량이 입력되지 않았습니다.", MsgBoxStyle.Information, "Chromaticity Rank 입력")
        '    Exit Sub
        'End If

        If MsgBox("저장 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = DBConnect1.BeginTransaction

        Try
            Dim sel_division As String = String.Empty
            Dim sel_use As String = String.Empty
            If CB_ModelDivision.Text = "개발" Then
                sel_division = "DEVE"
            ElseIf CB_ModelDivision.Text = "양산" Then
                sel_division = "MASS"
            End If
            If RB_USE.Checked Then
                sel_use = "Yes"
            ElseIf RB_NOT_USE.Checked Then
                sel_use = "No"
            End If
            If APP_SHEET_FILE.Text = "등록안됨" Then APP_SHEET_FILE.Text = String.Empty
            If FLOOR_PLAN_FILE.Text = "등록안됨" Then FLOOR_PLAN_FILE.Text = String.Empty
            If MARKING_FILE.Text = "등록안됨" Then MARKING_FILE.Text = String.Empty
            If OQC_FLOOR_PLAN_FILE.Text = "등록안됨" Then OQC_FLOOR_PLAN_FILE.Text = String.Empty
            If modify_model = False Then '신규모델 이라면
                strSQL += "insert into TB_MODEL_BLU_LIST(MODEL_RUN, DIVISION, MODEL_CODE, MODEL_NAME, APP_SHEET_DATE,"
                strSQL += " APP_SHEET_NAME, FLOOR_PLAN_DATE, FLOOR_PLAN_NAME, MARKING_DATE, MARKING_NAME,"
                strSQL += " OQC_FLOOR_PLAN_DATE, OQC_FLOOR_PLAN_NAME,"
                strSQL += " MODEL_NOTE, USER_ID, WRITE_DATE, PACKING_QTY, MODEL_PART_NO) values("
                strSQL += "'" & sel_use & "'"
                strSQL += ",'" & sel_division & "'"
                strSQL += ",'" & MODEL_CODE.Text & "'"
                strSQL += ",'" & MODEL_NAME.Text & "'"
                If APP_SHEET_FILE.Text = String.Empty Then
                    strSQL += ",null"
                    strSQL += ",null"
                Else
                    strSQL += ",'" & Format(APP_SHEET_DATE.Value, "yyyy-MM-dd") & "'"
                    strSQL += ",'" & APP_SHEET_FILE.Text & "'"
                End If
                If FLOOR_PLAN_FILE.Text = String.Empty Then
                    strSQL += ",null"
                    strSQL += ",null"
                Else
                    strSQL += ",'" & Format(FLOOR_PLAN_DATE.Value, "yyyy-MM-dd") & "'"
                    strSQL += ",'" & FLOOR_PLAN_FILE.Text & "'"
                End If
                If MARKING_FILE.Text = String.Empty Then
                    strSQL += ",null"
                    strSQL += ",null"
                Else
                    strSQL += ",'" & Format(MARKING_DATE.Value, "yyyy-MM-dd") & "'"
                    strSQL += ",'" & MARKING_FILE.Text & "'"
                End If
                If OQC_FLOOR_PLAN_FILE.Text = String.Empty Then
                    strSQL += ",null"
                    strSQL += ",null"
                Else
                    strSQL += ",'" & Format(OQC_FLOOR_PLAN_DATE.Value, "yyyy-MM-dd") & "'"
                    strSQL += ",'" & OQC_FLOOR_PLAN_FILE.Text & "'"
                End If
                strSQL += ",'" & MODEL_NOTE.Text & "'"
                strSQL += ",'" & Mainform.login_user & "'"
                strSQL += ",'" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "'"
                strSQL += ",'" & PACKING_QTY.Value & "'"
                strSQL += ",'" & MODEL_PART_NO.Text & "');"
                strSQL += "insert into TB_MODEL_BLU_BOM(MODEL_CODE, PKG_MAKER, PKG_PARTNO, PKG_QTY"
                strSQL += ", TH_MAKER, TH_PARTNO, TH_QTY, PCB_PARTNO) values("
                strSQL += "'" & MODEL_CODE.Text & "'"
                strSQL += ",'" & CB_PKG_Maker.Text & "'"
                strSQL += ",'" & PKG_PARTNO.Text & "'"
                strSQL += ",'" & PKG_QTY.Value & "'"
                strSQL += ",'" & CB_TH_Maker.Text & "'"
                strSQL += ",'" & TH_PARTNO.Text & "'"
                strSQL += ",'" & TH_QTY.Value & "'"
                strSQL += ",'" & PCB_PARTNO.Text & "');"
                strSQL += "insert into TB_MODEL_BLU_RANK(MODEL_CODE, INTENSITY, CHROMATICITY, RANK_ORDER) values("
                strSQL += "'" & MODEL_CODE.Text & "'"
                strSQL += ",'" & INTENSITY_RANK.Text & "'"
                strSQL += ",'" & CHROMATICITY_RANK.Text & "'"
                strSQL += ",'" & CB_RankOrder.Text & "');"
                'Dim total_strSQL As String = strSQL
                'strSQL += "insert into DB_HISTORY values("
                'strSQL += "'" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "'"
                'strSQL += ",'" & Replace(total_strSQL, "'", "!!/") & "'"
                'strSQL += ",'신규모델 정보 저장( " & MODEL_CODE.Text & " )'"
                'strSQL += ",'" & MMSLogin.USER_ID.Text & "');"
            Else '모델 수정이라면
                strSQL += "update TB_MODEL_BLU_LIST set MODEL_RUN = '" & sel_use & "'"
                strSQL += ", DIVISION = '" & sel_division & "'"
                strSQL += ", MODEL_NAME = '" & MODEL_NAME.Text & "'"
                If APP_SHEET_FILE.Text = String.Empty Then
                    strSQL += ", APP_SHEET_DATE = null"
                    strSQL += ", APP_SHEET_NAME = null"
                Else
                    strSQL += ", APP_SHEET_DATE = '" & Format(APP_SHEET_DATE.Value, "yyyy-MM-dd") & "'"
                    strSQL += ", APP_SHEET_NAME = '" & APP_SHEET_FILE.Text & "'"
                End If
                If FLOOR_PLAN_FILE.Text = String.Empty Then
                    strSQL += ", FLOOR_PLAN_DATE = null"
                    strSQL += ", FLOOR_PLAN_NAME = null"
                Else
                    strSQL += ", FLOOR_PLAN_DATE = '" & Format(FLOOR_PLAN_DATE.Value, "yyyy-MM-dd") & "'"
                    strSQL += ", FLOOR_PLAN_NAME = '" & FLOOR_PLAN_FILE.Text & "'"
                End If
                If MARKING_FILE.Text = String.Empty Then
                    strSQL += ", MARKING_DATE = null"
                    strSQL += ", MARKING_NAME = null"
                Else
                    strSQL += ", MARKING_DATE = '" & Format(MARKING_DATE.Value, "yyyy-MM-dd") & "'"
                    strSQL += ", MARKING_NAME = '" & MARKING_FILE.Text & "'"
                End If
                If OQC_FLOOR_PLAN_FILE.Text = String.Empty Then
                    strSQL += ", OQC_FLOOR_PLAN_DATE = null"
                    strSQL += ", OQC_FLOOR_PLAN_NAME = null"
                Else
                    strSQL += ", OQC_FLOOR_PLAN_DATE = '" & Format(OQC_FLOOR_PLAN_DATE.Value, "yyyy-MM-dd") & "'"
                    strSQL += ", OQC_FLOOR_PLAN_NAME = '" & OQC_FLOOR_PLAN_FILE.Text & "'"
                End If
                strSQL += ", MODEL_NOTE = '" & MODEL_NOTE.Text & "'"
                strSQL += ", MODIFY_DATE = '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "'"
                strSQL += ", MODIFY_ID = '" & Mainform.login_user & "'"
                strSQL += ", PACKING_QTY = '" & PACKING_QTY.Value & "'"
                strSQL += ", MODEL_PART_NO = '" & MODEL_PART_NO.Text & "'"
                strSQL += " where MODEL_CODE = '" & MODEL_CODE.Text & "';"
                strSQL += "update TB_MODEL_BLU_BOM set PKG_MAKER = '" & CB_PKG_Maker.Text & "'"
                strSQL += ", PKG_PARTNO = '" & PKG_PARTNO.Text & "'"
                strSQL += ", PKG_QTY = '" & PKG_QTY.Value & "'"
                strSQL += ", TH_MAKER = '" & CB_TH_Maker.Text & "'"
                strSQL += ", TH_PARTNO = '" & TH_PARTNO.Text & "'"
                strSQL += ", TH_QTY = '" & TH_QTY.Value & "'"
                strSQL += ", PCB_PARTNO = '" & PCB_PARTNO.Text & "'"
                strSQL += " where MODEL_CODE = '" & MODEL_CODE.Text & "';"
                strSQL += "update TB_MODEL_BLU_RANK set INTENSITY = '" & INTENSITY_RANK.Text & "'"
                strSQL += ", CHROMATICITY = '" & CHROMATICITY_RANK.Text & "'"
                strSQL += ", RANK_ORDER = '" & CB_RankOrder.Text & "'"
                strSQL += " where MODEL_CODE = '" & MODEL_CODE.Text & "';"
                'Dim total_strSQL As String = strSQL
                'strSQL += "insert into DB_HISTORY values("
                'strSQL += "'" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "'"
                'strSQL += ",'" & Replace(total_strSQL, "'", "!!/") & "'"
                'strSQL += ",'모델 정보 수정( " & MODEL_CODE.Text & " )'"
                'strSQL += ",'" & MMSLogin.USER_ID.Text & "');"
            End If

            sqlCmd = New MySqlCommand(strSQL, DBConnect1)
            sqlCmd.Transaction = sqlTran
            sqlCmd.ExecuteNonQuery()

            sqlTran.Commit()
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        Dim TEMP_FILE_Folder As String = Application.StartupPath & "\TEMP_FILE"

        If Not APP_SHEET_FILE.Text = String.Empty And app_sheet_new_upload = True Then
            '폴더(생성)
            MODEL_FILE_FOLDER(MODEL_CODE.Text, "APP_SHEET", Format(APP_SHEET_DATE.Value, "yyyyMMdd"))

            '파일(업로드)
            ftpFileUpload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/APP_SHEET/" & Format(APP_SHEET_DATE.Value, "yyyyMMdd"),
                              TEMP_FILE_Folder & "\" & APP_SHEET_FILE.Text)

            '임시저장 파일 삭제
            If My.Computer.FileSystem.FileExists(TEMP_FILE_Folder & "\" & APP_SHEET_FILE.Text) Then _
                My.Computer.FileSystem.DeleteFile(TEMP_FILE_Folder & "\" & APP_SHEET_FILE.Text)
        End If

        Application.DoEvents()

        If Not FLOOR_PLAN_FILE.Text = String.Empty And floor_plan_new_upload = True Then
            '폴더(생성)
            MODEL_FILE_FOLDER(MODEL_CODE.Text, "FLOOR_PLAN", Format(FLOOR_PLAN_DATE.Value, "yyyyMMdd"))

            '파일(업로드)
            ftpFileUpload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/FLOOR_PLAN/" & Format(FLOOR_PLAN_DATE.Value, "yyyyMMdd"),
                              TEMP_FILE_Folder & "\" & FLOOR_PLAN_FILE.Text)

            '임시저장 파일 삭제
            If My.Computer.FileSystem.FileExists(TEMP_FILE_Folder & "\" & FLOOR_PLAN_FILE.Text) Then _
                My.Computer.FileSystem.DeleteFile(TEMP_FILE_Folder & "\" & FLOOR_PLAN_FILE.Text)
        End If

        Application.DoEvents()

        If Not MARKING_FILE.Text = String.Empty And marking_new_upload = True Then
            '폴더(생성)
            MODEL_FILE_FOLDER(MODEL_CODE.Text, "MARKING", Format(MARKING_DATE.Value, "yyyyMMdd"))

            '파일(업로드)
            ftpFileUpload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/MARKING/" & Format(MARKING_DATE.Value, "yyyyMMdd"),
                              TEMP_FILE_Folder & "\" & MARKING_FILE.Text)

            '임시저장 파일 삭제
            If My.Computer.FileSystem.FileExists(TEMP_FILE_Folder & "\" & MARKING_FILE.Text) Then _
                My.Computer.FileSystem.DeleteFile(TEMP_FILE_Folder & "\" & MARKING_FILE.Text)
        End If

        Application.DoEvents()

        If Not OQC_FLOOR_PLAN_FILE.Text = String.Empty And oqc_floor_plan_new_upload = True Then
            '폴더(생성)
            MODEL_FILE_FOLDER(MODEL_CODE.Text, "OQC_FLOOR_PLAN", Format(OQC_FLOOR_PLAN_DATE.Value, "yyyyMMdd"))

            '파일(업로드)
            ftpFileUpload(ftpUrl & "/MODEL/" & MODEL_CODE.Text & "/OQC_FLOOR_PLAN/" & Format(OQC_FLOOR_PLAN_DATE.Value, "yyyyMMdd"),
                              TEMP_FILE_Folder & "\" & OQC_FLOOR_PLAN_FILE.Text)

            '임시저장 파일 삭제
            If My.Computer.FileSystem.FileExists(TEMP_FILE_Folder & "\" & OQC_FLOOR_PLAN_FILE.Text) Then _
                My.Computer.FileSystem.DeleteFile(TEMP_FILE_Folder & "\" & OQC_FLOOR_PLAN_FILE.Text)
        End If

        MsgBox("저장완료.", MsgBoxStyle.Information, form_Msgbox_String)

        BTN_SEARCH_Click(Nothing, Nothing)

    End Sub

    Private Sub CB_PKG_Maker_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles CB_PKG_Maker.DropDown

        Maker_DropDown_LIST(CB_PKG_Maker, "MC0003")

    End Sub

    Private Sub PKG_MAKER_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_PKG_Maker.SelectedIndexChanged

    End Sub

    Private Sub CB_TH_Maker_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles CB_TH_Maker.DropDown

        Maker_DropDown_LIST(CB_TH_Maker, "MC0003")

    End Sub

    Private Sub Maker_DropDown_LIST(ByVal control_name As ComboBox, ByVal division_code As String)

        DBConnect()

        'control_name.Items.Clear()

        Dim strSql As String = "select SUB_CODE_NAME from TB_SUB_CODE where MAIN_CODE = '" & division_code & "'"

        Dim sqlCmd As New MySqlCommand(strSql, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader
        Dim combo_list As String = String.Empty

        Do While sqlDR.Read
            Dim name_exist As Boolean = False
            For i = 0 To control_name.Items.Count - 1
                If sqlDR("SUB_CODE_NAME").ToString = control_name.Items(i).ToString Then
                    name_exist = True
                    Exit For
                End If
            Next
            If name_exist = False Then control_name.Items.Add(sqlDR("SUB_CODE_NAME"))
        Loop
        sqlDR.Close()

        DBClose()

    End Sub

    Private Sub CB_RankOrder_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_RankOrder.SelectedIndexChanged

        If Not CB_RankOrder.Text = String.Empty And
            Not INTENSITY_RANK.Text = String.Empty And
            Not CHROMATICITY_RANK.Text = String.Empty Then

            RANK_LIST_WRITE()

        End If

    End Sub

    Private Sub INTENSITY_RANK_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles INTENSITY_RANK.TextChanged

        If Not CB_RankOrder.Text = String.Empty And
            Not INTENSITY_RANK.Text = String.Empty And
            Not CHROMATICITY_RANK.Text = String.Empty Then

            RANK_LIST_WRITE()

        End If

    End Sub

    Private Sub CHROMATICITY_RANK_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHROMATICITY_RANK.TextChanged

        If Not CB_RankOrder.Text = String.Empty And
            Not INTENSITY_RANK.Text = String.Empty And
            Not CHROMATICITY_RANK.Text = String.Empty Then

            RANK_LIST_WRITE()

        End If

    End Sub

    Private Sub APP_SHEET_DATE_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles APP_SHEET_DATE.MouseHover

        APP_SHEET_DATE.Format = DateTimePickerFormat.Short

    End Sub

    Private Sub FLOOR_PLAN_DATE_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles FLOOR_PLAN_DATE.MouseHover

        FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Short

    End Sub

    Private Sub MARKING_DATE_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles MARKING_DATE.MouseHover

        MARKING_DATE.Format = DateTimePickerFormat.Short

    End Sub

    Private Sub OQC_FLOOR_PLAN_DATE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OQC_FLOOR_PLAN_DATE.MouseHover

        OQC_FLOOR_PLAN_DATE.Format = DateTimePickerFormat.Short

    End Sub
End Class
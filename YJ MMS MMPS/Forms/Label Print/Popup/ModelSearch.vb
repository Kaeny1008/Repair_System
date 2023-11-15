Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class ModelSearch

    Public findText As String

    Private Sub ModelSearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        GridSetting()

        If Not findText = String.Empty Then
            ModelList_Load(findText)
        End If

    End Sub

    Private Sub GridSetting()

        With GRID_ModelList
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 5
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            GRID_ModelList(0, 0) = "No"
            GRID_ModelList(0, 1) = "구분"
            GRID_ModelList(0, 2) = "MODEL Code"
            GRID_ModelList(0, 3) = "Model Name"
            GRID_ModelList(0, 4) = "비고"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter '비고란은 좌측 정렬
            .AutoSizeCols()
            .ShowCursor = True
        End With

    End Sub

    Private Sub ModelList_Load(ByVal modelName_Code As String)

        GRID_ModelList.Redraw = False

        DBConnect()

        GRID_ModelList.Rows.Count = 1

        Dim strSql As String = String.Empty

        strSql = "call USP_MODEL_SEARCH('" & modelName_Code & "')"

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

            Dim insert_String As String = row_num & vbTab &
                                          replace_Division & vbTab &
                                          sqlDR("MODEL_CODE") & vbTab &
                                          sqlDR("MODEL_NAME") & vbTab &
                                          sqlDR("MODEL_NOTE")
            GRID_ModelList.AddItem(insert_String)
        Loop
        sqlDR.Close()

        DBClose()

        GRID_ModelList.AutoSizeCols()

        GRID_ModelList.Redraw = True

    End Sub

    Private Sub BTN_ModelSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_ModelSearch.Click

        ModelList_Load(TB_ModelName.Text)

    End Sub

    Private Sub TB_ModelName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TB_ModelName.KeyDown

        If e.KeyCode = 13 Then
            ModelList_Load(TB_ModelName.Text)
        End If

    End Sub

    Private Sub GRID_ModelList_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GRID_ModelList.MouseDoubleClick

        If e.Button = Windows.Forms.MouseButtons.Left And GRID_ModelList.MouseRow > 0 Then
            GRID_ModelList.Row = GRID_ModelList.MouseRow
            BLULabelPrint.modelGubun = GRID_ModelList(GRID_ModelList.Row, 1)
            BLULabelPrint.TB_ModelCode.Text = GRID_ModelList(GRID_ModelList.Row, 2)
            BLULabelPrint.TB_ModelName.Text = GRID_ModelList(GRID_ModelList.Row, 3)
            Me.Dispose()
        End If

    End Sub
End Class
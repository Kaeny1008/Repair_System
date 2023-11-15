Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class AlarmHistory

    Public checkCode As String

    Private Sub AlarmHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GridSetting()

        DataLoad()

    End Sub

    Private Sub GridSetting()

        With Grid_historyList
            .Redraw = False
            .AllowEditing = False
            .AllowFiltering = True
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.Both
            .AllowDragging = AllowDraggingEnum.None
            .AllowResizing = AllowResizingEnum.None
            .KeyActionEnter = KeyActionEnum.MoveDown
            .KeyActionTab = KeyActionEnum.MoveAcross
            .Cols.Count = 6
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            Grid_historyList(0, 0) = "No"
            Grid_historyList(0, 1) = "Machine No."
            Grid_historyList(0, 2) = "Feeder No."
            Grid_historyList(0, 3) = "등록 Parts No."
            Grid_historyList(0, 4) = "확인 Parts No."
            Grid_historyList(0, 5) = "NG 사유"
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

    Private Sub DataLoad()

        Grid_historyList.Redraw = False
        Grid_historyList.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call USP_HISTORY(1"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",'" & checkCode & "'"
        strSQL += ",null"
        strSQL += ",null)"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert As String = String.Empty
            insert = Grid_historyList.Rows.Count &
                vbTab & sqlDR("MACHINE_NO") &
                vbTab & sqlDR("FEEDER_NO") &
                vbTab & sqlDR("ORG_PART_NO") &
                vbTab & sqlDR("CHECK_PART_NO") &
                vbTab & sqlDR("NG_RESULT")
            Grid_historyList.AddItem(insert)
        Loop
        sqlDR.Close()

        DBClose()

        Grid_historyList.AutoSizeCols()
        Grid_historyList.Redraw = True

    End Sub
End Class
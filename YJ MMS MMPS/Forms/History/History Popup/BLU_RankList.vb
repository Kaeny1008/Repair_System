Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class BLU_RankList

    Public checkCode As String

    Private Sub BLU_RankList_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GRID_Setting()

    End Sub

    Private Sub BLU_RankList_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        GRID_RankList.Redraw = False
        GRID_RankList.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call USP_HISTORY(7"
        strSQL += ",'2020-06-26 00:00:00'" 'startDate
        strSQL += ",'2020-06-26 00:00:00'" 'endDate
        strSQL += ",'" & checkCode & "'" 'customerName
        strSQL += ",null" 'modelName
        strSQL += ",null" 'factoryName
        strSQL += ",null" 'workLine
        strSQL += ",null" 'workSide
        strSQL += ",null" 'checkCode
        strSQL += ",null" 'checkResult
        strSQL += ",null)" 'firstSetting

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Dim allRank As String = String.Empty

        Do While sqlDR.Read
            allRank = sqlDR("AVAILABLE_RANK").ToString
        Loop
        sqlDR.Close()

        DBClose()

        If Not Trim(allRank) = String.Empty Then
            Dim splitRank() As String = allRank.Split(",")
            For i = 0 To UBound(splitRank)
                GRID_RankList.AddItem(GRID_RankList.Rows.Count - 1 & vbTab & splitRank(i))
            Next
        End If

        GRID_RankList.AutoSizeCols()
        GRID_RankList.Redraw = True

    End Sub

    Private Sub GRID_Setting()
        With GRID_RankList
            .Redraw = False
            .AllowEditing = False
            .AllowFiltering = True
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.Both
            .AllowDragging = AllowDraggingEnum.None
            .AllowResizing = AllowResizingEnum.None
            .KeyActionEnter = KeyActionEnum.MoveDown
            .KeyActionTab = KeyActionEnum.MoveAcross
            .Cols.Count = 2
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            GRID_RankList(0, 0) = "No"
            GRID_RankList(0, 1) = "사용 가능 RANK"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter '비고란은 좌측 정렬
            .ExtendLastCol = True '마지막 열 확장
            .AutoSizeCols()
            .ShowCursor = True
            .Cols.Frozen = 0
            .FocusRect = FocusRectEnum.Light
            .Redraw = True
        End With
    End Sub
End Class
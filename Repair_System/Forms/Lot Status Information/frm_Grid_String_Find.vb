Imports C1.Win.C1FlexGrid

Public Class frm_Grid_String_Find

    Public first_row As Integer
    Public first_col As Integer
    Public selGrid

    Private Sub frm_Grid_String_Find_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox1.SelectAll()
        TextBox1.Focus()

    End Sub

    Private Sub frm_Grid_String_Find_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Select Case e.KeyCode
            Case 27
                btn_Close_Click(Nothing, Nothing)
            Case 13
                btn_Find_Click(Nothing, Nothing)
        End Select

    End Sub

    Private Sub btn_Close_Click(sender As Object, e As EventArgs) Handles btn_Close.Click

        Me.Close()

    End Sub

    Private Sub btn_Find_Click(sender As Object, e As EventArgs) Handles btn_Find.Click

        Console.WriteLine(first_row & ", " & first_col)

        Dim fullMatch As Boolean = False
        Dim findRow As Integer = 0
        Dim findCol As Integer = 0

        If CheckBox1.Checked Then fullMatch = True

        With selGrid
            For j = first_col + 1 To .Cols.Count - 1
                If fullMatch = True Then
                    If selGrid(first_row, j).ToString.ToUpper.Equals(TextBox1.Text.ToUpper) Then
                        findRow = first_row
                        findCol = j
                        first_col = j
                        Exit For
                    End If
                Else
                    If selGrid(first_row, j).ToString.ToUpper Like TextBox1.Text.ToUpper & "*" Then
                        findRow = first_row
                        findCol = j
                        first_col = j
                        Exit For
                    End If
                End If
            Next

            If findRow = 0 Or findCol = 0 Then
                For i = first_row + 1 To .Rows.Count - 1
                    For j = 1 To .Cols.Count - 1
                        If fullMatch = True Then
                            If selGrid(i, j).ToString.ToUpper.Equals(TextBox1.Text.ToUpper) Then
                                findRow = i
                                findCol = j
                                first_row = i
                                first_col = j
                                Exit For
                            End If
                        Else
                            If selGrid(i, j).ToString.ToUpper Like TextBox1.Text.ToUpper & "*" Then
                                findRow = i
                                findCol = j
                                first_row = i
                                first_col = j
                                Exit For
                            End If
                        End If
                    Next
                    If findRow > 0 Or findCol > 0 Then Exit For
                Next
            End If
        End With

        If findRow = 0 Then
            If MsgBox("찾는 내용없음." & vbCrLf & "처음부터 재검색 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "찾기") = MsgBoxResult.No Then Exit Sub
            first_col = 0
            first_row = 0
            retryFind()
        Else
            selGrid.Select(findRow, findCol, True)
        End If

    End Sub
    Private Sub retryFind()

        Dim fullMatch As Boolean = False
        Dim findRow As Integer = 0
        Dim findCol As Integer = 0

        If CheckBox1.Checked Then fullMatch = True

        With selGrid
            For i = first_row + 1 To .Rows.Count - 1
                For j = 1 To .Cols.Count - 1
                    If fullMatch = True Then
                        If selGrid(i, j).ToString.ToUpper.Equals(TextBox1.Text.ToUpper) Then
                            findRow = i
                            findCol = j
                            first_row = i
                            first_col = j
                            Exit For
                        End If
                    Else
                        If selGrid(i, j).ToString.ToUpper Like TextBox1.Text.ToUpper & "*" Then
                            findRow = i
                            findCol = j
                            first_row = i
                            first_col = j
                            Exit For
                        End If
                    End If
                Next
                If findRow > 0 Or findCol > 0 Then Exit For
            Next
        End With

        If findRow = 0 Then
            MsgBox("찾는 내용없음.", MsgBoxStyle.Information, "찾기")
        Else
            selGrid.Select(findRow, findCol, True)
        End If

    End Sub

    Private Sub frm_Grid_String_Find_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus

        TextBox1.SelectAll()
        TextBox1.Focus()

    End Sub


End Class
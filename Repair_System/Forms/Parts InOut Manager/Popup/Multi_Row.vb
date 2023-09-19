Public Class Multi_Row

    Private Sub Multi_Row_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TB_Row_Count_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TB_Row_Count.KeyDown

        Select Case e.KeyCode
            Case 96 To 105, 48 To 57 '숫자키
            Case 13 'Enter
                BTN_Insert_Click(Nothing, Nothing)
            Case 8 'Backspace
            Case 46 'Delete
            Case 37 To 40 'Cusor
            Case Else
                e.SuppressKeyPress = True
        End Select

    End Sub

    Private Sub BTN_Insert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Insert.Click

        'If TB_Row_Count.Text = String.Empty Then Exit Sub
        'frm_Module_In_Add.temp_inCode = Format(Now, "yyyyMMddHHmmss") - 1
        'Dim i As Integer
        'For i = 1 To TB_Row_Count.Text
        '    frm_Module_In_Add.Row_ADD_Click("다중 행 추가", Nothing)
        'Next
        'Me.Dispose()
        'frm_Module_In_Add.temp_inCode = 0

    End Sub
End Class
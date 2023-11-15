Public Class frm_LoadingImage
    Private Sub frm_LoadingImage_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Timer1.Interval = 10
        'Timer1.Enabled = True

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If formClose = True Then
            Me.DialogResult = DialogResult.None
            Me.Dispose()
        End If

    End Sub
End Class
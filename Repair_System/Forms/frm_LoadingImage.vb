Public Class frm_LoadingImage
    Private Sub frm_LoadingImage_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Timer1.Interval = 1
        Timer1.Enabled = True

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        'Console.WriteLine(formClose)
        If formClose = True Then
            Console.WriteLine("(Loading Form) Finished...")
            Me.DialogResult = DialogResult.OK
            Me.Dispose()
        End If

    End Sub
End Class
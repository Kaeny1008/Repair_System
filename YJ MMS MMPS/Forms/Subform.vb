Public Class Subform

    Private Sub Subform_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

    End Sub

    Dim ddw_old_color As Color = Color.SaddleBrown
    Dim history_old_color As Color = Color.SteelBlue
    Dim exit_old_color As Color = Color.DarkMagenta
    Dim codemanager_old_color As Color = Color.DarkGoldenrod
    Dim modelmanager_old_color As Color = Color.BurlyWood
    Dim labelPrint_old_color As Color = Color.Peru
    Dim history2_old_color As Color = Color.SandyBrown

    'Dim old_color As Color

    Private Sub Btn_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btn_DDW.MouseMove,
                                                                                            Btn_Exit.MouseMove,
                                                                                            Btn_History.MouseMove,
                                                                                            Btn_CodeManager.MouseMove,
                                                                                            Btn_ModelManager.MouseMove,
                                                                                            BTN_LabelPrint.MouseMove,
                                                                                            BTN_HistoryBLU.MouseMove

        Dim abcd As Button = CType(sender, Button)

        abcd.BackColor = Color.AliceBlue
        abcd.ForeColor = Color.Black

    End Sub

    Private Sub Btn_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btn_DDW.MouseLeave,
                                                                                            Btn_History.MouseLeave,
                                                                                            Btn_ModelManager.MouseLeave,
                                                                                            Btn_Exit.MouseLeave,
                                                                                            Btn_CodeManager.MouseLeave,
                                                                                            BTN_LabelPrint.MouseLeave,
                                                                                            BTN_HistoryBLU.MouseLeave

        Dim abcd As Button = CType(sender, Button)

        If abcd.Name = Btn_DDW.Name Then
            abcd.BackColor = ddw_old_color
        ElseIf abcd.Name = Btn_History.Name Then
            abcd.BackColor = history_old_color
        ElseIf abcd.Name = Btn_Exit.Name Then
            abcd.BackColor = exit_old_color
        ElseIf abcd.Name = Btn_CodeManager.Name Then
            abcd.BackColor = codemanager_old_color
        ElseIf abcd.Name = Btn_ModelManager.Name Then
            abcd.BackColor = modelmanager_old_color
        ElseIf abcd.Name = BTN_LabelPrint.Name Then
            abcd.BackColor = labelPrint_old_color
        ElseIf abcd.Name = BTN_HistoryBLU.Name Then
            abcd.BackColor = history2_old_color
        End If

        abcd.ForeColor = Color.White

    End Sub

    Private Sub Btn_Exit_Click(sender As Object, e As EventArgs) Handles Btn_Exit.Click

        Application.Exit()

    End Sub

    Private Sub Btn_DDW_Click(sender As Object, e As EventArgs) Handles Btn_DDW.Click

        Mainform.form_name = DeviceData

        DeviceData.MdiParent = Mainform
        If Not DeviceData.Visible Then DeviceData.Show()
        DeviceData.Focus()

    End Sub

    Private Sub Btn_CodeManager_Click(sender As Object, e As EventArgs) Handles Btn_CodeManager.Click

        Mainform.form_name = CodeManager

        CodeManager.MdiParent = Mainform
        If Not CodeManager.Visible Then CodeManager.Show()
        CodeManager.Focus()

    End Sub

    Private Sub Btn_History_Click(sender As Object, e As EventArgs) Handles Btn_History.Click

        Mainform.form_name = HistoryForm

        HistoryForm.MdiParent = Mainform
        If Not HistoryForm.Visible Then HistoryForm.Show()
        HistoryForm.Focus()

    End Sub

    Private Sub BTN_connection_Click(sender As Object, e As EventArgs) Handles BTN_connection.Click

        SQL_Connection.ShowDialog()

    End Sub

    Private Sub BTN_uploadPG_Click(sender As Object, e As EventArgs) Handles BTN_uploadPG.Click

        '사용금지 현재 리페어 시스템 디비와 연결되지 않음.

        MsgBox("지원되지 않습니다.", MsgBoxStyle.Information, "YJ MMS MMPS")
        'PG_UPLOAD.ShowDialog()

    End Sub

    Private Sub Btn_MouseHover(sender As Object, e As MouseEventArgs) Handles Btn_History.MouseMove, Btn_DDW.MouseMove, Btn_ModelManager.MouseMove, Btn_Exit.MouseMove, Btn_CodeManager.MouseMove, BTN_LabelPrint.MouseMove, BTN_HistoryBLU.MouseMove

    End Sub

    Private Sub Btn_ModelManager_Click(sender As Object, e As EventArgs) Handles Btn_ModelManager.Click

        Mainform.form_name = BluModelManager

        BluModelManager.MdiParent = Mainform
        If Not BluModelManager.Visible Then BluModelManager.Show()
        BluModelManager.Focus()

    End Sub

    Private Sub BTN_LabelPrint_Click(sender As Object, e As EventArgs) Handles BTN_LabelPrint.Click

        Mainform.form_name = BLULabelPrint

        BLULabelPrint.MdiParent = Mainform
        If Not BLULabelPrint.Visible Then BLULabelPrint.Show()
        BLULabelPrint.Focus()

    End Sub

    Private Sub BTN_HistoryBLU_Click(sender As Object, e As EventArgs) Handles BTN_HistoryBLU.Click

        Mainform.form_name = HistoryBLUForm

        HistoryBLUForm.MdiParent = Mainform
        If Not HistoryBLUForm.Visible Then HistoryBLUForm.Show()
        HistoryBLUForm.Focus()

    End Sub
End Class

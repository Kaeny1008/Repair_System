Public Class Subform
    Private Sub Subform_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

    End Sub

    Dim ddw_old_color As Color = Color.SaddleBrown
    Dim apc_old_color As Color = Color.DarkGoldenrod
    Dim pc_old_color As Color = Color.Green
    Dim pc_old_color2 As Color = Color.DarkOliveGreen
    Dim history_old_color As Color = Color.SteelBlue
    Dim exit_old_color As Color = Color.DarkMagenta
    Dim inspection_old_color As Color = Color.CadetBlue
    Dim cal_old_color As Color = Color.DarkBlue

    Private Sub Btn_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTN_Code_Manager.MouseMove,
                                                                                            BTN_Asset_ADD.MouseMove,
                                                                                            BTN_Close.MouseMove,
                                                                                            BTN_Asset_History.MouseMove,
                                                                                            BTN_Asset_Manager.MouseMove,
                                                                                            BTN_Asset_Manager2.MouseMove,
                                                                                            BTN_Asset_Inspection.MouseMove,
                                                                                            BTN_Cal_Object.MouseMove

        Dim abcd As Button = CType(sender, Button)

        abcd.BackColor = Color.AliceBlue
        abcd.ForeColor = Color.Black

    End Sub

    Private Sub Btn_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTN_Code_Manager.MouseLeave,
                                                                                            BTN_Asset_ADD.MouseLeave,
                                                                                            BTN_Close.MouseLeave,
                                                                                            BTN_Asset_History.MouseLeave,
                                                                                            BTN_Asset_Manager.MouseLeave,
                                                                                            BTN_Asset_Manager2.MouseLeave,
                                                                                            BTN_Asset_Inspection.MouseLeave,
                                                                                            BTN_Cal_Object.MouseLeave

        Dim abcd As Button = CType(sender, Button)

        If abcd.Name = BTN_Code_Manager.Name Then
            abcd.BackColor = ddw_old_color
        ElseIf abcd.Name = BTN_Asset_ADD.Name Then
            abcd.BackColor = apc_old_color
        ElseIf abcd.Name = BTN_Asset_Manager.Name Then
            abcd.BackColor = pc_old_color
        ElseIf abcd.Name = BTN_Asset_History.Name Then
            abcd.BackColor = history_old_color
        ElseIf abcd.Name = BTN_Close.Name Then
            abcd.BackColor = exit_old_color
        ElseIf abcd.Name = BTN_Asset_Manager2.Name Then
            abcd.BackColor = pc_old_color2
        ElseIf abcd.Name = BTN_Asset_Inspection.Name Then
            abcd.BackColor = inspection_old_color
        ElseIf abcd.Name = BTN_cal_object.Name Then
            abcd.BackColor = cal_old_color
        End If

        abcd.ForeColor = Color.White

    End Sub

    Private Sub BTN_Close_Click(sender As Object, e As EventArgs) Handles BTN_Close.Click

        Application.Exit()

    End Sub

    Public Sub BTN_Code_Manager_Click(sender As Object, e As EventArgs) Handles BTN_Code_Manager.Click

        CODE_MANAGER.MdiParent = Mainform
        If Not CODE_MANAGER.Visible Then CODE_MANAGER.Show()
        CODE_MANAGER.Focus()

        Mainform.form_name = CODE_MANAGER

        For i = 0 To Mainform.ToolStrip.Items.Count - 1
            If Mainform.ToolStrip.Items(i).Text = BTN_Code_Manager.Text Then
                Exit Sub
            End If
        Next

        Mainform.ToolStrip.Items.Add("Code 관리", YJ_MMS_Asset.My.Resources.Resources.product_sales_report_128)

    End Sub

    Public Sub BTN_Asset_ADD_Click(sender As Object, e As EventArgs) Handles BTN_Asset_ADD.Click

        ASSET_REGISTRATION.MdiParent = Mainform
        If Not ASSET_REGISTRATION.Visible Then ASSET_REGISTRATION.Show()
        ASSET_REGISTRATION.Focus()

        Mainform.form_name = ASSET_REGISTRATION

        For i = 0 To Mainform.ToolStrip.Items.Count - 1
            If Mainform.ToolStrip.Items(i).Text = BTN_Asset_ADD.Text Then
                Exit Sub
            End If
        Next

        Mainform.ToolStrip.Items.Add("자산등록", YJ_MMS_Asset.My.Resources.Resources.Open_Folder_Add)

    End Sub

    Public Sub BTN_Asset_Manager_Click(sender As Object, e As EventArgs) Handles BTN_Asset_Manager.Click

        ASSET_HISTORY.MdiParent = Mainform
        If Not ASSET_HISTORY.Visible Then ASSET_HISTORY.Show()
        ASSET_HISTORY.Focus()

        Mainform.form_name = ASSET_HISTORY

        For i = 0 To Mainform.ToolStrip.Items.Count - 1
            If Mainform.ToolStrip.Items(i).Text = BTN_Asset_Manager.Text Then
                Exit Sub
            End If
        Next

        Mainform.ToolStrip.Items.Add("자산관리(개별)", YJ_MMS_Asset.My.Resources.Resources.web_management)

    End Sub

    Public Sub BTN_Asset_History_Click(sender As Object, e As EventArgs) Handles BTN_Asset_History.Click

        ASSET_LIST.MdiParent = Mainform
        If Not ASSET_LIST.Visible Then ASSET_LIST.Show()
        ASSET_LIST.Focus()

        Mainform.form_name = ASSET_LIST

        For i = 0 To Mainform.ToolStrip.Items.Count - 1
            If Mainform.ToolStrip.Items(i).Text = BTN_Asset_History.Text Then
                Exit Sub
            End If
        Next

        Mainform.ToolStrip.Items.Add("자산현황", YJ_MMS_Asset.My.Resources.Resources.Star_Full_128)

    End Sub

    Public Sub BTN_Asset_Manager2_Click(sender As Object, e As EventArgs) Handles BTN_Asset_Manager2.Click

        ASSET_HISTORY2.MdiParent = Mainform
        If Not ASSET_HISTORY2.Visible Then ASSET_HISTORY2.Show()
        ASSET_HISTORY2.Focus()

        Mainform.form_name = ASSET_HISTORY2

        For i = 0 To Mainform.ToolStrip.Items.Count - 1
            If Mainform.ToolStrip.Items(i).Text = BTN_Asset_Manager2.Text Then
                Exit Sub
            End If
        Next

        Mainform.ToolStrip.Items.Add("자산관리(일괄)", YJ_MMS_Asset.My.Resources.Resources.Globe_Download_128)

    End Sub

    Private Sub Btn_MouseHover(sender As Object, e As MouseEventArgs) Handles BTN_Code_Manager.MouseMove, BTN_Close.MouseMove, BTN_Asset_Manager2.MouseMove, BTN_Asset_Manager.MouseMove, BTN_Asset_History.MouseMove, BTN_Asset_ADD.MouseMove, BTN_Asset_Inspection.MouseMove, BTN_Cal_Object.MouseMove

    End Sub

    Public Sub BTN_Asset_Inspection_Click(sender As Object, e As EventArgs) Handles BTN_Asset_Inspection.Click

        ASSET_INSPECTION.MdiParent = Mainform
        If Not ASSET_INSPECTION.Visible Then ASSET_INSPECTION.Show()
        ASSET_INSPECTION.Focus()

        Mainform.form_name = ASSET_INSPECTION

        For i = 0 To Mainform.ToolStrip.Items.Count - 1
            If Mainform.ToolStrip.Items(i).Text = BTN_Asset_Inspection.Text Then
                Exit Sub
            End If
        Next

        Mainform.ToolStrip.Items.Add("재물조사", YJ_MMS_Asset.My.Resources.Resources.ordering_128)

    End Sub

    Private Sub BTN_connection_Click(sender As Object, e As EventArgs) Handles BTN_connection.Click

        SQL_Connection.ShowDialog()

    End Sub

    Private Sub BTN_uploadPG_Click(sender As Object, e As EventArgs) Handles BTN_uploadPG.Click

        PG_UPLOAD.ShowDialog()

    End Sub

    Public Sub BTN_Cal_Object_Click(sender As Object, e As EventArgs) Handles BTN_Cal_Object.Click

        CALIBRATION_OBJECT.MdiParent = Mainform
        If Not CALIBRATION_OBJECT.Visible Then CALIBRATION_OBJECT.Show()
        CALIBRATION_OBJECT.Focus()

        Mainform.form_name = CALIBRATION_OBJECT

        For i = 0 To Mainform.ToolStrip.Items.Count - 1
            If Mainform.ToolStrip.Items(i).Text = BTN_Cal_Object.Text Then
                Exit Sub
            End If
        Next

        Mainform.ToolStrip.Items.Add("검교정 관리", YJ_MMS_Asset.My.Resources.Resources.Process_Info_128)

    End Sub
End Class
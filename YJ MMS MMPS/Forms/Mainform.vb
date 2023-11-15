Imports System.ComponentModel
Imports MySql.Data.MySqlClient

Public Class Mainform

    Public login_user As String = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS\Login", "USER ID", "ADMIN")
    Public login_name As String = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS\Login", "USER Name", "Admin")
    Dim form_msgbox_string As String = "오삽방지 시스템"
    Public form_name As Form

    Private Sub Mainform_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Subform.MdiParent = Me
        Subform.Show()

        MMPS_DB_Registry()

        Me.Text = "Yujin Manufacturing Management System : 오삽방지 시스템 ( Ver. " &
            System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() & " ) "

        UpdatePG_Run()

    End Sub

    Private Sub UpdatePG_Run()

        Dim pgClose As Boolean = False
        If UBound(Diagnostics.Process.GetProcessesByName("UpdateChecker")) < 0 Then
            Try
                Shell(Application.StartupPath & "\UpdateChecker.exe", AppWinStyle.Hide)
            Catch ex As Exception
                If MsgBox("업데이트 확인 프로그램을 실행 할 수 없습니다." & vbCrLf &
                          "무시 후 계속 진행하시겠습니까?" & vbCrLf &
                          "( " & Application.StartupPath & "\UpdateChecker.exe" & " )",
                            MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.No Then
                    pgClose = True
                End If
            End Try
        End If

        If pgClose = True Then Me.Close()

    End Sub

    Private Sub BTN_Home_Click(sender As Object, e As EventArgs) Handles BTN_Home.Click

        Subform.Focus()
        form_name = Nothing

    End Sub

    Private Sub ToolStrip_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip.ItemClicked

        'If e.ClickedItem.Text = "Code 관리" Then
        '    Subform.BTN_Code_Manager_Click(Nothing, Nothing)
        'ElseIf e.ClickedItem.Text = "자산등록" Then
        '    Subform.BTN_Asset_ADD_Click(Nothing, Nothing)
        'ElseIf e.ClickedItem.Text = "자산관리(개별)" Then
        '    Subform.BTN_Asset_Manager_Click(Nothing, Nothing)
        'ElseIf e.ClickedItem.Text = "자산현황" Then
        '    Subform.BTN_Asset_History_Click(Nothing, Nothing)
        'ElseIf e.ClickedItem.Text = "자산관리(일괄)" Then
        '    Subform.BTN_Asset_Manager2_Click(Nothing, Nothing)
        'ElseIf e.ClickedItem.Text = "재물조사" Then
        '    Subform.BTN_Asset_Inspection_Click(Nothing, Nothing)
        'ElseIf e.ClickedItem.Text = "검교정 관리" Then
        '    Subform.BTN_Cal_Object_Click(Nothing, Nothing)
        'End If

    End Sub

    Private Sub Mainform_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        If IO.Directory.Exists(Application.StartupPath & "\TEMP_FILE") Then
            IO.Directory.Delete(Application.StartupPath & "\TEMP_FILE", True)
        End If

        Dim ucProcess() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("UpdateChecker")

        '아래 목록은 YJ 프로그램이 실행된 상태인지 확인 후 하나라도 켜져 있다면 업데이트체커 프로그램을 끄지 않도록 한다.
        '현재 내자신(프로그램)을 제외하고 실행되었는지 확인 후 없다면 업데이트체커 프로그램을 종료시킨다.
        Dim mmpsProcess() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("YJ MMS MMPS")
        Dim mmsProcess() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("YJ MMS MMMS")
        Dim labelProcess() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("YJ MMS Label Printer")

        If ucProcess.Length > 0 Then
            If mmsProcess.Length = 0 And labelProcess.Length = 0 Then
                ucProcess(0).Kill()
            End If
        End If
        'If MsgBox("프로그램을 종료 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.Yes Then Application.Exit()

    End Sub

    Private Sub Btn_DDW_Click(sender As Object, e As EventArgs) Handles Btn_DDW.Click

        form_name = DeviceData

        DeviceData.MdiParent = Me
        If Not DeviceData.Visible Then DeviceData.Show()
        DeviceData.Focus()

    End Sub

    Private Sub Btn_CodeManager_Click(sender As Object, e As EventArgs) Handles Btn_CodeManager.Click

        form_name = CodeManager

        CodeManager.MdiParent = Me
        If Not CodeManager.Visible Then CodeManager.Show()
        CodeManager.Focus()

    End Sub

    Private Sub BTN_Form_Close_Click(sender As Object, e As EventArgs) Handles BTN_Form_Close.Click

        If IsNothing(form_name) Then Exit Sub

        form_name.Dispose()

    End Sub

    Private Sub Btn_History_Click(sender As Object, e As EventArgs) Handles Btn_History.Click

        form_name = HistoryForm

        HistoryForm.MdiParent = Me
        If Not HistoryForm.Visible Then HistoryForm.Show()
        HistoryForm.Focus()

    End Sub

    Private Sub Btn_ModelManager_Click(sender As Object, e As EventArgs) Handles Btn_ModelManager.Click

        form_name = BluModelManager

        BluModelManager.MdiParent = Me
        If Not BluModelManager.Visible Then BluModelManager.Show()
        BluModelManager.Focus()

    End Sub

    Private Sub BTN_LabelPrint_Click(sender As Object, e As EventArgs) Handles BTN_LabelPrint.Click

        form_name = BLULabelPrint

        BLULabelPrint.MdiParent = Me
        If Not BLULabelPrint.Visible Then BLULabelPrint.Show()
        BLULabelPrint.Focus()

    End Sub

    Private Sub BTN_HistoryBLU_Click(sender As Object, e As EventArgs) Handles BTN_HistoryBLU.Click

        form_name = HistoryBLUForm

        HistoryBLUForm.MdiParent = Me
        If Not HistoryBLUForm.Visible Then HistoryBLUForm.Show()
        HistoryBLUForm.Focus()

    End Sub
End Class

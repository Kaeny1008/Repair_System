Imports System.ComponentModel
Imports System.Threading

Public Class frm_Main

    Public form_name As Form

    Private Sub frm_Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "Repair System (Build " &
            System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() &
            ")"

        lb_LoginID.Text = lb_LoginID.Text & loginID

        frm_Sub.MdiParent = Me
        frm_Sub.Show()

        lb_Status.Text = String.Empty

        Me.WindowState = FormWindowState.Maximized

        UpdatePG_Run()

        'If Not loginID = "ADMIN" Then
        'menu_Monthly_Production_Report.Visible = False
        'End If

        Timer1.Interval = 100
        Timer1.Enabled = True

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
                            MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Update Checker") = MsgBoxResult.No Then
                    pgClose = True
                End If
            End Try
        End If

        If pgClose = True Then Me.Close()

    End Sub

    Private Sub btnHome_Click(sender As Object, e As EventArgs) Handles btnHome.Click

        frm_Sub.Focus()
        form_name = Nothing

    End Sub

    Private Sub btnSQLConn_Click(sender As Object, e As EventArgs) Handles btn_SQLConn.Click

        frm_SQL_Conn.ShowDialog()

    End Sub

    Private Sub btnCodeManager_Click(sender As Object, e As EventArgs) Handles btn_CodeManager.Click

        frm_code_manager.MdiParent = Me
        If Not frm_code_manager.Visible Then frm_code_manager.Show()
        frm_code_manager.Focus()

    End Sub

    Private Sub btn_PGUpdate_Click(sender As Object, e As EventArgs) Handles btn_PGUpdate.Click

        frm_PG_Upload.MdiParent = Me
        If Not frm_PG_Upload.Visible Then frm_PG_Upload.Show()
        frm_PG_Upload.Focus()

    End Sub

    Private Sub btn_OMSFileInfo_Click(sender As Object, e As EventArgs) Handles btn_OMSFileInfo.Click

        frm_oms_file_data.MdiParent = Me
        If Not frm_oms_file_data.Visible Then frm_oms_file_data.Show()
        frm_oms_file_data.Focus()

    End Sub

    Private Sub btn_Bucket_Combination_Click(sender As Object, e As EventArgs) Handles btn_Bucket_Combination.Click

        frm_Bucket_RCD_PMIC.MdiParent = Me
        If Not frm_Bucket_RCD_PMIC.Visible Then frm_Bucket_RCD_PMIC.Show()
        frm_Bucket_RCD_PMIC.Focus()

    End Sub

    Private Sub btn_Module_In_Add_Click(sender As Object, e As EventArgs) Handles btn_Module_In_Add.Click

        frm_Module_In_Add.MdiParent = Me
        If Not frm_Module_In_Add.Visible Then frm_Module_In_Add.Show()
        frm_Module_In_Add.Focus()

    End Sub

    Private Sub btn_OMS_WIP_Data_Click(sender As Object, e As EventArgs) Handles btn_OMS_WIP_Data.Click

        frm_OMS_WIP_DATA.MdiParent = Me
        If Not frm_OMS_WIP_DATA.Visible Then frm_OMS_WIP_DATA.Show()
        frm_OMS_WIP_DATA.Focus()

    End Sub

    Private Sub btn_IQC_Print_Click(sender As Object, e As EventArgs) Handles btn_IQC_Print.Click

        frm_IQC_Print.MdiParent = Me
        If Not frm_IQC_Print.Visible Then frm_IQC_Print.Show()
        frm_IQC_Print.Focus()

    End Sub

    Private Sub btn_Material_In_Add_Click(sender As Object, e As EventArgs) Handles btn_Material_In_Add.Click

        frm_Material_In_Add.MdiParent = Me
        If Not frm_Material_In_Add.Visible Then frm_Material_In_Add.Show()
        frm_Material_In_Add.Focus()

    End Sub

    Private Sub btn_IQC_Standby_List_Click(sender As Object, e As EventArgs) Handles btn_IQC_Standby_List.Click

        frm_IQC_Standby_List.MdiParent = Me
        If Not frm_IQC_Standby_List.Visible Then frm_IQC_Standby_List.Show()
        frm_IQC_Standby_List.Focus()

    End Sub

    Private Sub btn_Label_RePrint_Click(sender As Object, e As EventArgs) Handles btn_Label_RePrint.Click

        frm_IQC_RePrint.MdiParent = Me
        If Not frm_IQC_RePrint.Visible Then frm_IQC_RePrint.Show()
        frm_IQC_RePrint.Focus()

    End Sub

    Private Sub btn_Work_Baking_Click(sender As Object, e As EventArgs) Handles btn_Work_Baking.Click

        frm_Work_Baking.MdiParent = Me
        If Not frm_Work_Baking.Visible Then frm_Work_Baking.Show()
        frm_Work_Baking.Focus()

    End Sub

    Private Sub btn_Baking_History_Click(sender As Object, e As EventArgs) Handles btn_Baking_History.Click

        frm_Baking_History.MdiParent = Me
        If Not frm_Baking_History.Visible Then frm_Baking_History.Show()
        frm_Baking_History.Focus()

    End Sub

    Private Sub btn_PartsKitting_Click(sender As Object, e As EventArgs) Handles btn_PartsKitting.Click

        frm_Parts_Kitting.MdiParent = Me
        If Not frm_Parts_Kitting.Visible Then frm_Parts_Kitting.Show()
        frm_Parts_Kitting.Focus()

    End Sub

    Private Sub frm_Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        If MsgBox("정말 종료하시겠습니까?", vbYesNo + vbQuestion, "Repair System") = vbYes Then '메세지박스 띄우기
            End '종료
        Else
            e.Cancel = True
        End If

        If IO.Directory.Exists(Application.StartupPath & "\TEMP_FILE") Then
            IO.Directory.Delete(Application.StartupPath & "\TEMP_FILE", True)
        End If

        Dim ucProcess() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("UpdateChecker")

        '아래 목록은 YJ 프로그램이 실행된 상태인지 확인 후 하나라도 켜져 있다면 업데이트체커 프로그램을 끄지 않도록 한다.
        '현재 내자신(프로그램)을 제외하고 실행되었는지 확인 후 없다면 업데이트체커 프로그램을 종료시킨다.
        Dim repairSystemProcess() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("Repair System")

        If ucProcess.Length > 0 Then
            If repairSystemProcess.Length = 0 Then
                ucProcess(0).Kill()
            End If
        End If
        'If MsgBox("프로그램을 종료 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_msgbox_string) = MsgBoxResult.Yes Then Application.Exit()

    End Sub

    Private Sub btn_LabelPrinter_Click(sender As Object, e As EventArgs) Handles btn_LabelPrinter.Click

        frm_LabelPrinter.MdiParent = Me
        If Not frm_LabelPrinter.Visible Then frm_LabelPrinter.Show()
        frm_LabelPrinter.Focus()

    End Sub

    Private Sub btn_WorkSite_Moving_Click(sender As Object, e As EventArgs) Handles btn_WorkSite_Moving.Click

        frm_WorkSite_Moving.MdiParent = Me
        If Not frm_WorkSite_Moving.Visible Then frm_WorkSite_Moving.Show()
        frm_WorkSite_Moving.Focus()

    End Sub

    Private Sub btn_Module_Out_Add_Click(sender As Object, e As EventArgs) Handles btn_Module_Out_Add.Click

        frm_Module_Out_Add.MdiParent = Me
        If Not frm_Module_Out_Add.Visible Then frm_Module_Out_Add.Show()
        frm_Module_Out_Add.Focus()

    End Sub

    Private Sub btn_OMS_Ship_Data_Click(sender As Object, e As EventArgs) Handles btn_OMS_Ship_Data.Click

        frm_OMS_Ship_Data.MdiParent = Me
        If Not frm_OMS_Ship_Data.Visible Then frm_OMS_Ship_Data.Show()
        frm_OMS_Ship_Data.Focus()

    End Sub

    Private Sub btn_Reject_IC_In_Click(sender As Object, e As EventArgs) Handles btn_Reject_IC_In.Click

        frm_Reject_IC_In.MdiParent = Me
        If Not frm_Reject_IC_In.Visible Then frm_Reject_IC_In.Show()
        frm_Reject_IC_In.Focus()

    End Sub

    Private Sub btn_Lot_Total_Information_Click(sender As Object, e As EventArgs) Handles btn_Lot_Total_Information.Click

        frm_Lot_Total_Information.MdiParent = Me
        If Not frm_Lot_Total_Information.Visible Then frm_Lot_Total_Information.Show()
        frm_Lot_Total_Information.Focus()

    End Sub

    Private Sub btn_SMT_Working_History_Click(sender As Object, e As EventArgs) Handles btn_SMT_Working_History.Click

        frm_SMT_Working_History.MdiParent = Me
        If Not frm_SMT_Working_History.Visible Then frm_SMT_Working_History.Show()
        frm_SMT_Working_History.Focus()

    End Sub

    Private Sub btn_MisCheck_Result_Click(sender As Object, e As EventArgs) Handles btn_MisCheck_Result.Click

        frm_MisCheck_Result.MdiParent = Me
        If Not frm_MisCheck_Result.Visible Then frm_MisCheck_Result.Show()
        frm_MisCheck_Result.Focus()

    End Sub

    Private Sub btn_Material_Total_Information_Click(sender As Object, e As EventArgs) Handles btn_Material_Lot_Information.Click

        frm_Material_LotNo_Information.MdiParent = Me
        If Not frm_Material_LotNo_Information.Visible Then frm_Material_LotNo_Information.Show()
        frm_Material_LotNo_Information.Focus()

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click

        frm_Material_PartNo_Information.MdiParent = Me
        If Not frm_Material_PartNo_Information.Visible Then frm_Material_PartNo_Information.Show()
        frm_Material_PartNo_Information.Focus()

    End Sub

    Private Sub btn_Material_Total_Information_Click_1(sender As Object, e As EventArgs) Handles btn_Material_Total_Information.Click

        frm_Material_Total_Information.MdiParent = Me
        If Not frm_Material_Total_Information.Visible Then frm_Material_Total_Information.Show()
        frm_Material_Total_Information.Focus()

    End Sub

    Private Sub btn_Reject_IC_Ship_Click(sender As Object, e As EventArgs) Handles btn_Reject_IC_Ship.Click

        frm_Reject_IC_Ship.MdiParent = Me
        If Not frm_Reject_IC_Ship.Visible Then frm_Reject_IC_Ship.Show()
        frm_Reject_IC_Ship.Focus()

    End Sub

    Private Sub btn_OMS_Monthly_Data_Save_Click(sender As Object, e As EventArgs) Handles btn_OMS_Monthly_Data_Save.Click

        frm_Monthly_Production_Data.MdiParent = Me
        If Not frm_Monthly_Production_Data.Visible Then frm_Monthly_Production_Data.Show()
        frm_Monthly_Production_Data.Focus()

    End Sub

    Private Sub btn_OMS_Monthly_Data_Click(sender As Object, e As EventArgs) Handles btn_OMS_Monthly_Data.Click

        frm_Monthly_Production_Report.MdiParent = Me
        If Not frm_Monthly_Production_Report.Visible Then frm_Monthly_Production_Report.Show()
        frm_Monthly_Production_Report.Focus()

    End Sub

    Private Sub btn_Lot_Residence_Time_Click(sender As Object, e As EventArgs) Handles btn_Lot_Residence_Time.Click

        frm_Lot_Residence_Time.MdiParent = Me
        If Not frm_Lot_Residence_Time.Visible Then frm_Lot_Residence_Time.Show()
        frm_Lot_Residence_Time.Focus()

    End Sub
End Class

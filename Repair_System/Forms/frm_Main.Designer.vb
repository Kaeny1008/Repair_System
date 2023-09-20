<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_Main
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub


    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Main))
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.lb_LoginID = New System.Windows.Forms.ToolStripStatusLabel()
        Me.pgbMain = New System.Windows.Forms.ToolStripProgressBar()
        Me.lb_Status = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStrip = New System.Windows.Forms.ToolStrip()
        Me.btnHome = New System.Windows.Forms.ToolStripButton()
        Me.btnSetting = New System.Windows.Forms.ToolStripDropDownButton()
        Me.btn_SQLConn = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_CodeManager = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_UserManager = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_PrinterSetting = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_PGUpdate = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_OMSFileInfo = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_BasicInfo = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_Bucket_Combination = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_LabelPrinter = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_PartsManager = New System.Windows.Forms.ToolStripDropDownButton()
        Me.btn_Material_In_Add = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_Module_In_Add = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Lot_Total_Information = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Material_Lot_Information = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_Material_Total_Information = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_OMS_WIP_Data = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_incoming_inspection = New System.Windows.Forms.ToolStripDropDownButton()
        Me.btn_IQC_Print = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_Label_RePrint = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_IQC_Standby_List = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripDropDownButton1 = New System.Windows.Forms.ToolStripDropDownButton()
        Me.btn_Work_Baking = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_Baking_History = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripDropDownButton2 = New System.Windows.Forms.ToolStripDropDownButton()
        Me.btn_PartsKitting = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_WorkSite_Moving = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripDropDownButton()
        Me.btn_Module_Out_Add = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator10 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_OMS_Ship_Data = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripDropDownButton3 = New System.Windows.Forms.ToolStripDropDownButton()
        Me.btn_Reject_IC_In = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Reject_IC_Ship = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripDropDownButton4 = New System.Windows.Forms.ToolStripDropDownButton()
        Me.btn_SMT_Working_History = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_MisCheck_Result = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.menu_Monthly_Production_Report = New System.Windows.Forms.ToolStripDropDownButton()
        Me.btn_OMS_Monthly_Data_Save = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_OMS_Monthly_Data = New System.Windows.Forms.ToolStripMenuItem()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.btn_Lot_Residence_Time = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip.SuspendLayout()
        Me.ToolStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Padding = New System.Windows.Forms.Padding(7, 2, 0, 2)
        Me.MenuStrip.Size = New System.Drawing.Size(1123, 24)
        Me.MenuStrip.TabIndex = 11
        Me.MenuStrip.Text = "MenuStrip"
        Me.MenuStrip.Visible = False
        '
        'StatusStrip
        '
        Me.StatusStrip.BackColor = System.Drawing.SystemColors.Control
        Me.StatusStrip.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lb_LoginID, Me.pgbMain, Me.lb_Status})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 713)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Padding = New System.Windows.Forms.Padding(1, 0, 16, 0)
        Me.StatusStrip.Size = New System.Drawing.Size(1123, 35)
        Me.StatusStrip.TabIndex = 14
        Me.StatusStrip.Text = "StatusStrip"
        '
        'lb_LoginID
        '
        Me.lb_LoginID.AutoSize = False
        Me.lb_LoginID.Name = "lb_LoginID"
        Me.lb_LoginID.Size = New System.Drawing.Size(300, 30)
        Me.lb_LoginID.Text = "Login ID : "
        Me.lb_LoginID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pgbMain
        '
        Me.pgbMain.Name = "pgbMain"
        Me.pgbMain.Size = New System.Drawing.Size(100, 29)
        Me.pgbMain.Visible = False
        '
        'lb_Status
        '
        Me.lb_Status.AutoSize = False
        Me.lb_Status.Name = "lb_Status"
        Me.lb_Status.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lb_Status.Size = New System.Drawing.Size(806, 30)
        Me.lb_Status.Spring = True
        Me.lb_Status.Text = "lb_Status"
        Me.lb_Status.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ToolStrip
        '
        Me.ToolStrip.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnHome, Me.btnSetting, Me.ToolStripSeparator1, Me.btn_PartsManager, Me.btn_incoming_inspection, Me.ToolStripDropDownButton1, Me.ToolStripDropDownButton2, Me.ToolStripButton1, Me.ToolStripDropDownButton3, Me.ToolStripSeparator6, Me.ToolStripDropDownButton4, Me.ToolStripSeparator9, Me.menu_Monthly_Production_Report})
        Me.ToolStrip.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip.Name = "ToolStrip"
        Me.ToolStrip.Size = New System.Drawing.Size(1123, 31)
        Me.ToolStrip.TabIndex = 13
        Me.ToolStrip.Text = "ToolStrip"
        '
        'btnHome
        '
        Me.btnHome.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btnHome.Image = Global.Repair_System.My.Resources.Resources.home
        Me.btnHome.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnHome.Name = "btnHome"
        Me.btnHome.Size = New System.Drawing.Size(80, 28)
        Me.btnHome.Text = "Home"
        '
        'btnSetting
        '
        Me.btnSetting.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_SQLConn, Me.btn_CodeManager, Me.btn_UserManager, Me.btn_PrinterSetting, Me.btn_PGUpdate, Me.ToolStripSeparator3, Me.btn_OMSFileInfo, Me.btn_BasicInfo, Me.btn_LabelPrinter})
        Me.btnSetting.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.btnSetting.Image = Global.Repair_System.My.Resources.Resources.sitemap
        Me.btnSetting.ImageTransparentColor = System.Drawing.Color.Black
        Me.btnSetting.Name = "btnSetting"
        Me.btnSetting.Size = New System.Drawing.Size(76, 28)
        Me.btnSetting.Text = "설정"
        '
        'btn_SQLConn
        '
        Me.btn_SQLConn.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_SQLConn.Name = "btn_SQLConn"
        Me.btn_SQLConn.Size = New System.Drawing.Size(206, 22)
        Me.btn_SQLConn.Text = "SQL Connection Setting"
        '
        'btn_CodeManager
        '
        Me.btn_CodeManager.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_CodeManager.Name = "btn_CodeManager"
        Me.btn_CodeManager.Size = New System.Drawing.Size(206, 22)
        Me.btn_CodeManager.Text = "Code 관리"
        '
        'btn_UserManager
        '
        Me.btn_UserManager.Enabled = False
        Me.btn_UserManager.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_UserManager.Name = "btn_UserManager"
        Me.btn_UserManager.Size = New System.Drawing.Size(206, 22)
        Me.btn_UserManager.Text = "사용자 관리"
        Me.btn_UserManager.Visible = False
        '
        'btn_PrinterSetting
        '
        Me.btn_PrinterSetting.Enabled = False
        Me.btn_PrinterSetting.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_PrinterSetting.Name = "btn_PrinterSetting"
        Me.btn_PrinterSetting.Size = New System.Drawing.Size(206, 22)
        Me.btn_PrinterSetting.Text = "라벨 프린터 설정"
        Me.btn_PrinterSetting.Visible = False
        '
        'btn_PGUpdate
        '
        Me.btn_PGUpdate.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_PGUpdate.Name = "btn_PGUpdate"
        Me.btn_PGUpdate.Size = New System.Drawing.Size(206, 22)
        Me.btn_PGUpdate.Text = "프로그램 파일 Upload"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(203, 6)
        '
        'btn_OMSFileInfo
        '
        Me.btn_OMSFileInfo.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_OMSFileInfo.Name = "btn_OMSFileInfo"
        Me.btn_OMSFileInfo.Size = New System.Drawing.Size(206, 22)
        Me.btn_OMSFileInfo.Text = "OMS 파일 기초정보"
        '
        'btn_BasicInfo
        '
        Me.btn_BasicInfo.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Bucket_Combination})
        Me.btn_BasicInfo.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_BasicInfo.Name = "btn_BasicInfo"
        Me.btn_BasicInfo.Size = New System.Drawing.Size(206, 22)
        Me.btn_BasicInfo.Text = "작업별 기초정보"
        '
        'btn_Bucket_Combination
        '
        Me.btn_Bucket_Combination.Name = "btn_Bucket_Combination"
        Me.btn_Bucket_Combination.Size = New System.Drawing.Size(271, 22)
        Me.btn_Bucket_Combination.Text = "Bucket별 PMIC/RCD 조합 및 Vendor"
        '
        'btn_LabelPrinter
        '
        Me.btn_LabelPrinter.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_LabelPrinter.Name = "btn_LabelPrinter"
        Me.btn_LabelPrinter.Size = New System.Drawing.Size(206, 22)
        Me.btn_LabelPrinter.Text = "각종 라벨 발행"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 31)
        '
        'btn_PartsManager
        '
        Me.btn_PartsManager.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btn_PartsManager.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Material_In_Add, Me.btn_Module_In_Add, Me.ToolStripSeparator5, Me.btn_Lot_Residence_Time, Me.btn_Lot_Total_Information, Me.ToolStripSeparator7, Me.btn_Material_Lot_Information, Me.ToolStripMenuItem1, Me.btn_Material_Total_Information, Me.ToolStripSeparator2, Me.btn_OMS_WIP_Data})
        Me.btn_PartsManager.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.btn_PartsManager.Image = CType(resources.GetObject("btn_PartsManager.Image"), System.Drawing.Image)
        Me.btn_PartsManager.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_PartsManager.Name = "btn_PartsManager"
        Me.btn_PartsManager.Size = New System.Drawing.Size(99, 28)
        Me.btn_PartsManager.Text = "입고 / 재공"
        '
        'btn_Material_In_Add
        '
        Me.btn_Material_In_Add.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_Material_In_Add.Name = "btn_Material_In_Add"
        Me.btn_Material_In_Add.Size = New System.Drawing.Size(219, 22)
        Me.btn_Material_In_Add.Text = "자재 입고 등록"
        '
        'btn_Module_In_Add
        '
        Me.btn_Module_In_Add.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_Module_In_Add.Name = "btn_Module_In_Add"
        Me.btn_Module_In_Add.Size = New System.Drawing.Size(219, 22)
        Me.btn_Module_In_Add.Text = "Module / Comp 입고 등록"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(216, 6)
        '
        'btn_Lot_Total_Information
        '
        Me.btn_Lot_Total_Information.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_Lot_Total_Information.Name = "btn_Lot_Total_Information"
        Me.btn_Lot_Total_Information.Size = New System.Drawing.Size(219, 22)
        Me.btn_Lot_Total_Information.Text = "Module Lot 정보보기"
        '
        'ToolStripSeparator7
        '
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        Me.ToolStripSeparator7.Size = New System.Drawing.Size(216, 6)
        '
        'btn_Material_Lot_Information
        '
        Me.btn_Material_Lot_Information.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_Material_Lot_Information.Name = "btn_Material_Lot_Information"
        Me.btn_Material_Lot_Information.Size = New System.Drawing.Size(219, 22)
        Me.btn_Material_Lot_Information.Text = "자재 Lot별 정보보기"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(219, 22)
        Me.ToolStripMenuItem1.Text = "자재 Part No.별 정보보기"
        '
        'btn_Material_Total_Information
        '
        Me.btn_Material_Total_Information.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_Material_Total_Information.Name = "btn_Material_Total_Information"
        Me.btn_Material_Total_Information.Size = New System.Drawing.Size(219, 22)
        Me.btn_Material_Total_Information.Text = "자재 종합정보"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(216, 6)
        '
        'btn_OMS_WIP_Data
        '
        Me.btn_OMS_WIP_Data.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_OMS_WIP_Data.Name = "btn_OMS_WIP_Data"
        Me.btn_OMS_WIP_Data.Size = New System.Drawing.Size(219, 22)
        Me.btn_OMS_WIP_Data.Text = "OMS 재공데이터 생성"
        '
        'btn_incoming_inspection
        '
        Me.btn_incoming_inspection.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btn_incoming_inspection.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_IQC_Print, Me.btn_Label_RePrint, Me.ToolStripSeparator4, Me.btn_IQC_Standby_List})
        Me.btn_incoming_inspection.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.btn_incoming_inspection.Image = CType(resources.GetObject("btn_incoming_inspection.Image"), System.Drawing.Image)
        Me.btn_incoming_inspection.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_incoming_inspection.Name = "btn_incoming_inspection"
        Me.btn_incoming_inspection.Size = New System.Drawing.Size(82, 28)
        Me.btn_incoming_inspection.Text = "수입검사"
        '
        'btn_IQC_Print
        '
        Me.btn_IQC_Print.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_IQC_Print.Name = "btn_IQC_Print"
        Me.btn_IQC_Print.Size = New System.Drawing.Size(212, 22)
        Me.btn_IQC_Print.Text = "검사등록 및 라벨 발행"
        '
        'btn_Label_RePrint
        '
        Me.btn_Label_RePrint.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_Label_RePrint.Name = "btn_Label_RePrint"
        Me.btn_Label_RePrint.Size = New System.Drawing.Size(212, 22)
        Me.btn_Label_RePrint.Text = "라벨 추가 발행 및 재발행"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(209, 6)
        '
        'btn_IQC_Standby_List
        '
        Me.btn_IQC_Standby_List.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_IQC_Standby_List.Name = "btn_IQC_Standby_List"
        Me.btn_IQC_Standby_List.Size = New System.Drawing.Size(212, 22)
        Me.btn_IQC_Standby_List.Text = "수입검사 대기 리스트"
        '
        'ToolStripDropDownButton1
        '
        Me.ToolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripDropDownButton1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Work_Baking, Me.btn_Baking_History})
        Me.ToolStripDropDownButton1.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.ToolStripDropDownButton1.Image = CType(resources.GetObject("ToolStripDropDownButton1.Image"), System.Drawing.Image)
        Me.ToolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton1.Name = "ToolStripDropDownButton1"
        Me.ToolStripDropDownButton1.Size = New System.Drawing.Size(70, 28)
        Me.ToolStripDropDownButton1.Text = "Baking"
        '
        'btn_Work_Baking
        '
        Me.btn_Work_Baking.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_Work_Baking.Name = "btn_Work_Baking"
        Me.btn_Work_Baking.Size = New System.Drawing.Size(205, 22)
        Me.btn_Work_Baking.Text = "Baking 작업 투입 / 완료"
        '
        'btn_Baking_History
        '
        Me.btn_Baking_History.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btn_Baking_History.Name = "btn_Baking_History"
        Me.btn_Baking_History.Size = New System.Drawing.Size(205, 22)
        Me.btn_Baking_History.Text = "Baking 작업현황"
        '
        'ToolStripDropDownButton2
        '
        Me.ToolStripDropDownButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripDropDownButton2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_PartsKitting, Me.btn_WorkSite_Moving})
        Me.ToolStripDropDownButton2.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.ToolStripDropDownButton2.Image = CType(resources.GetObject("ToolStripDropDownButton2.Image"), System.Drawing.Image)
        Me.ToolStripDropDownButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton2.Name = "ToolStripDropDownButton2"
        Me.ToolStripDropDownButton2.Size = New System.Drawing.Size(117, 28)
        Me.ToolStripDropDownButton2.Text = "Kitting / 출고"
        '
        'btn_PartsKitting
        '
        Me.btn_PartsKitting.Enabled = False
        Me.btn_PartsKitting.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_PartsKitting.Name = "btn_PartsKitting"
        Me.btn_PartsKitting.Size = New System.Drawing.Size(152, 22)
        Me.btn_PartsKitting.Text = "Kitting 등록"
        '
        'btn_WorkSite_Moving
        '
        Me.btn_WorkSite_Moving.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_WorkSite_Moving.Name = "btn_WorkSite_Moving"
        Me.btn_WorkSite_Moving.Size = New System.Drawing.Size(152, 22)
        Me.btn_WorkSite_Moving.Text = "현장출고 등록"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Module_Out_Add, Me.ToolStripSeparator10, Me.btn_OMS_Ship_Data})
        Me.ToolStripButton1.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(139, 28)
        Me.ToolStripButton1.Text = "출고(유진->우익)"
        '
        'btn_Module_Out_Add
        '
        Me.btn_Module_Out_Add.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_Module_Out_Add.Name = "btn_Module_Out_Add"
        Me.btn_Module_Out_Add.Size = New System.Drawing.Size(239, 22)
        Me.btn_Module_Out_Add.Text = "출고 등록(전표생성)"
        '
        'ToolStripSeparator10
        '
        Me.ToolStripSeparator10.Name = "ToolStripSeparator10"
        Me.ToolStripSeparator10.Size = New System.Drawing.Size(236, 6)
        '
        'btn_OMS_Ship_Data
        '
        Me.btn_OMS_Ship_Data.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_OMS_Ship_Data.Name = "btn_OMS_Ship_Data"
        Me.btn_OMS_Ship_Data.Size = New System.Drawing.Size(239, 22)
        Me.btn_OMS_Ship_Data.Text = "OMS 반제품 출고 데이터 생성"
        '
        'ToolStripDropDownButton3
        '
        Me.ToolStripDropDownButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripDropDownButton3.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Reject_IC_In, Me.ToolStripSeparator8, Me.btn_Reject_IC_Ship})
        Me.ToolStripDropDownButton3.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.ToolStripDropDownButton3.Image = CType(resources.GetObject("ToolStripDropDownButton3.Image"), System.Drawing.Image)
        Me.ToolStripDropDownButton3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton3.Name = "ToolStripDropDownButton3"
        Me.ToolStripDropDownButton3.Size = New System.Drawing.Size(82, 28)
        Me.ToolStripDropDownButton3.Text = "폐기자재"
        '
        'btn_Reject_IC_In
        '
        Me.btn_Reject_IC_In.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_Reject_IC_In.Name = "btn_Reject_IC_In"
        Me.btn_Reject_IC_In.Size = New System.Drawing.Size(152, 22)
        Me.btn_Reject_IC_In.Text = "폐기자재 회수"
        '
        'ToolStripSeparator8
        '
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        Me.ToolStripSeparator8.Size = New System.Drawing.Size(149, 6)
        '
        'btn_Reject_IC_Ship
        '
        Me.btn_Reject_IC_Ship.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_Reject_IC_Ship.Name = "btn_Reject_IC_Ship"
        Me.btn_Reject_IC_Ship.Size = New System.Drawing.Size(152, 22)
        Me.btn_Reject_IC_Ship.Text = "폐기자재 반납"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(6, 31)
        '
        'ToolStripDropDownButton4
        '
        Me.ToolStripDropDownButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripDropDownButton4.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_SMT_Working_History, Me.btn_MisCheck_Result})
        Me.ToolStripDropDownButton4.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.ToolStripDropDownButton4.Image = CType(resources.GetObject("ToolStripDropDownButton4.Image"), System.Drawing.Image)
        Me.ToolStripDropDownButton4.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton4.Name = "ToolStripDropDownButton4"
        Me.ToolStripDropDownButton4.Size = New System.Drawing.Size(88, 28)
        Me.ToolStripDropDownButton4.Text = "SMT 관리"
        '
        'btn_SMT_Working_History
        '
        Me.btn_SMT_Working_History.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_SMT_Working_History.Name = "btn_SMT_Working_History"
        Me.btn_SMT_Working_History.Size = New System.Drawing.Size(208, 22)
        Me.btn_SMT_Working_History.Text = "SMT 생산현황"
        '
        'btn_MisCheck_Result
        '
        Me.btn_MisCheck_Result.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_MisCheck_Result.Name = "btn_MisCheck_Result"
        Me.btn_MisCheck_Result.Size = New System.Drawing.Size(208, 22)
        Me.btn_MisCheck_Result.Text = "자재사용 현황(오삽검사)"
        '
        'ToolStripSeparator9
        '
        Me.ToolStripSeparator9.Name = "ToolStripSeparator9"
        Me.ToolStripSeparator9.Size = New System.Drawing.Size(6, 31)
        '
        'menu_Monthly_Production_Report
        '
        Me.menu_Monthly_Production_Report.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.menu_Monthly_Production_Report.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_OMS_Monthly_Data_Save, Me.btn_OMS_Monthly_Data})
        Me.menu_Monthly_Production_Report.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.menu_Monthly_Production_Report.Image = CType(resources.GetObject("menu_Monthly_Production_Report.Image"), System.Drawing.Image)
        Me.menu_Monthly_Production_Report.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.menu_Monthly_Production_Report.Name = "menu_Monthly_Production_Report"
        Me.menu_Monthly_Production_Report.Size = New System.Drawing.Size(82, 28)
        Me.menu_Monthly_Production_Report.Text = "생산월보"
        '
        'btn_OMS_Monthly_Data_Save
        '
        Me.btn_OMS_Monthly_Data_Save.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_OMS_Monthly_Data_Save.Name = "btn_OMS_Monthly_Data_Save"
        Me.btn_OMS_Monthly_Data_Save.Size = New System.Drawing.Size(183, 22)
        Me.btn_OMS_Monthly_Data_Save.Text = "OMS 생산월보 저장"
        '
        'btn_OMS_Monthly_Data
        '
        Me.btn_OMS_Monthly_Data.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_OMS_Monthly_Data.Name = "btn_OMS_Monthly_Data"
        Me.btn_OMS_Monthly_Data.Size = New System.Drawing.Size(183, 22)
        Me.btn_OMS_Monthly_Data.Text = "생산월보 확인"
        '
        'btn_Lot_Residence_Time
        '
        Me.btn_Lot_Residence_Time.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.btn_Lot_Residence_Time.Name = "btn_Lot_Residence_Time"
        Me.btn_Lot_Residence_Time.Size = New System.Drawing.Size(219, 22)
        Me.btn_Lot_Residence_Time.Text = "Lot 정체일(현장출고 후)"
        '
        'frm_Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1123, 748)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.ToolStrip)
        Me.Controls.Add(Me.MenuStrip)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "frm_Main"
        Me.Text = "Repair System (Build 230707)"
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ToolStrip.ResumeLayout(False)
        Me.ToolStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip As MenuStrip
    Friend WithEvents StatusStrip As StatusStrip
    Friend WithEvents lb_LoginID As ToolStripStatusLabel
    Friend WithEvents lb_Status As ToolStripStatusLabel
    Friend WithEvents ToolStrip As ToolStrip
    Friend WithEvents btnHome As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents btn_PartsManager As ToolStripDropDownButton
    Friend WithEvents btn_Module_In_Add As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents btn_OMS_WIP_Data As ToolStripMenuItem
    Friend WithEvents btn_incoming_inspection As ToolStripDropDownButton
    Friend WithEvents btn_IQC_Print As ToolStripMenuItem
    Friend WithEvents btn_Material_In_Add As ToolStripMenuItem
    Friend WithEvents btn_IQC_Standby_List As ToolStripMenuItem
    Friend WithEvents btn_Label_RePrint As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents ToolStripDropDownButton1 As ToolStripDropDownButton
    Friend WithEvents btn_Work_Baking As ToolStripMenuItem
    Friend WithEvents btn_Baking_History As ToolStripMenuItem
    Friend WithEvents ToolStripDropDownButton2 As ToolStripDropDownButton
    Friend WithEvents btn_PartsKitting As ToolStripMenuItem
    Friend WithEvents btn_WorkSite_Moving As ToolStripMenuItem
    Friend WithEvents ToolStripButton1 As ToolStripDropDownButton
    Friend WithEvents btn_Module_Out_Add As ToolStripMenuItem
    Friend WithEvents btn_OMS_Ship_Data As ToolStripMenuItem
    Friend WithEvents ToolStripDropDownButton3 As ToolStripDropDownButton
    Friend WithEvents btn_Reject_IC_In As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents btn_Lot_Total_Information As ToolStripMenuItem
    Friend WithEvents ToolStripDropDownButton4 As ToolStripDropDownButton
    Friend WithEvents btn_SMT_Working_History As ToolStripMenuItem
    Friend WithEvents btn_MisCheck_Result As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator6 As ToolStripSeparator
    Friend WithEvents btn_Material_Lot_Information As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator7 As ToolStripSeparator
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents btn_Material_Total_Information As ToolStripMenuItem
    Friend WithEvents btnSetting As ToolStripDropDownButton
    Friend WithEvents btn_SQLConn As ToolStripMenuItem
    Friend WithEvents btn_CodeManager As ToolStripMenuItem
    Friend WithEvents btn_UserManager As ToolStripMenuItem
    Friend WithEvents btn_PrinterSetting As ToolStripMenuItem
    Friend WithEvents btn_PGUpdate As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents btn_OMSFileInfo As ToolStripMenuItem
    Friend WithEvents btn_BasicInfo As ToolStripMenuItem
    Friend WithEvents btn_Bucket_Combination As ToolStripMenuItem
    Friend WithEvents btn_LabelPrinter As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator8 As ToolStripSeparator
    Friend WithEvents btn_Reject_IC_Ship As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator9 As ToolStripSeparator
    Friend WithEvents menu_Monthly_Production_Report As ToolStripDropDownButton
    Friend WithEvents btn_OMS_Monthly_Data_Save As ToolStripMenuItem
    Friend WithEvents btn_OMS_Monthly_Data As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator10 As ToolStripSeparator
    Friend WithEvents pgbMain As ToolStripProgressBar
    Friend WithEvents Timer1 As Timer
    Friend WithEvents btn_Lot_Residence_Time As ToolStripMenuItem
End Class

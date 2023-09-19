<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_IQC_RePrint
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
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_IQC_RePrint))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.BTN_SavePrint = New System.Windows.Forms.Button()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.tb_Process_No = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.tb_Process_Char = New System.Windows.Forms.TextBox()
        Me.tb_Requester = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cb_Request_Reason = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tb_lotNo = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cb_Process = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.btn_TestLabel = New System.Windows.Forms.Button()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.cb_PrinterList = New System.Windows.Forms.ComboBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.CB_Print_Use = New System.Windows.Forms.CheckBox()
        Me.TB_StopBits = New System.Windows.Forms.TextBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.TB_Parity = New System.Windows.Forms.TextBox()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.cb_Cable = New System.Windows.Forms.ComboBox()
        Me.TB_DataBits = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.TB_TOP_Loc = New System.Windows.Forms.TextBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.TB_BaudRate = New System.Windows.Forms.TextBox()
        Me.TB_LEFT_Loc = New System.Windows.Forms.TextBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.TB_Port = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.grid_History = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.tb_Lot_Option = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.tb_total_print_qty = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.tb_After_RCD_Marking = New System.Windows.Forms.TextBox()
        Me.tb_Before_RCD_Marking = New System.Windows.Forms.TextBox()
        Me.tb_After_PMIC_Marking = New System.Windows.Forms.TextBox()
        Me.tb_Before_PMIC_Marking = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tb_After_RCD = New System.Windows.Forms.TextBox()
        Me.tb_Before_RCD = New System.Windows.Forms.TextBox()
        Me.tb_After_PMIC = New System.Windows.Forms.TextBox()
        Me.tb_Before_PMIC = New System.Windows.Forms.TextBox()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.tb_After_Bucket = New System.Windows.Forms.TextBox()
        Me.tb_before_Bucket = New System.Windows.Forms.TextBox()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.tb_Print_YJ_No = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tb_Print_Lot_No = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.btn_Parts_Save = New System.Windows.Forms.Button()
        Me.tb_Parts_Requester = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cb_Parts_Reason = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.tb_Parts_LotNo = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.tb_Parts_Load_TopMarking = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.tb_Parts_Load_YJNo = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.tb_Parts_Load_LotNo = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.tb_Parts_Load_PartNo = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Toolstrip1 = New System.Windows.Forms.ToolStrip()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        Me.BTN_Setting_Open = New System.Windows.Forms.ToolStripButton()
        Me.GroupBox1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel5.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.grid_History, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.Toolstrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.White
        Me.GroupBox1.Controls.Add(Me.Label21)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.BTN_SavePrint)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Controls.Add(Me.tb_Process_No)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.tb_Process_Char)
        Me.GroupBox1.Controls.Add(Me.tb_Requester)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.cb_Request_Reason)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.tb_lotNo)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cb_Process)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(6, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(717, 147)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "발행 정보"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label21.Location = New System.Drawing.Point(394, 28)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(113, 12)
        Me.Label21.TabIndex = 13
        Me.Label21.Text = "<-- 입력 후 꼭 엔터"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.RadioButton2.Location = New System.Drawing.Point(459, 106)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(75, 16)
        Me.RadioButton2.TabIndex = 12
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "추가 발행"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'BTN_SavePrint
        '
        Me.BTN_SavePrint.BackColor = System.Drawing.Color.Red
        Me.BTN_SavePrint.Font = New System.Drawing.Font("굴림", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_SavePrint.Image = Global.Repair_System.My.Resources.Resources.barcode
        Me.BTN_SavePrint.Location = New System.Drawing.Point(551, 22)
        Me.BTN_SavePrint.Name = "BTN_SavePrint"
        Me.BTN_SavePrint.Size = New System.Drawing.Size(153, 108)
        Me.BTN_SavePrint.TabIndex = 1
        Me.BTN_SavePrint.Text = "저장 및 발행"
        Me.BTN_SavePrint.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BTN_SavePrint.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.RadioButton1.Location = New System.Drawing.Point(459, 84)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(59, 16)
        Me.RadioButton1.TabIndex = 11
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "재발행"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'tb_Process_No
        '
        Me.tb_Process_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Process_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Process_No.Location = New System.Drawing.Point(149, 93)
        Me.tb_Process_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Process_No.Name = "tb_Process_No"
        Me.tb_Process_No.Size = New System.Drawing.Size(242, 21)
        Me.tb_Process_No.TabIndex = 10
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(15, 93)
        Me.Label8.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(134, 21)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "공정번호(재발행)"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Process_Char
        '
        Me.tb_Process_Char.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Process_Char.Enabled = False
        Me.tb_Process_Char.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Process_Char.Location = New System.Drawing.Point(351, 70)
        Me.tb_Process_Char.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Process_Char.Name = "tb_Process_Char"
        Me.tb_Process_Char.Size = New System.Drawing.Size(40, 21)
        Me.tb_Process_Char.TabIndex = 6
        Me.tb_Process_Char.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_Requester
        '
        Me.tb_Requester.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Requester.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Requester.Location = New System.Drawing.Point(149, 47)
        Me.tb_Requester.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Requester.Name = "tb_Requester"
        Me.tb_Requester.Size = New System.Drawing.Size(242, 21)
        Me.tb_Requester.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(15, 47)
        Me.Label4.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(134, 21)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "요청자"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cb_Request_Reason
        '
        Me.cb_Request_Reason.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.cb_Request_Reason.FormattingEnabled = True
        Me.cb_Request_Reason.Items.AddRange(New Object() {"라벨 훼손", "라벨 분실", "추가 (불량)"})
        Me.cb_Request_Reason.Location = New System.Drawing.Point(149, 116)
        Me.cb_Request_Reason.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.cb_Request_Reason.Name = "cb_Request_Reason"
        Me.cb_Request_Reason.Size = New System.Drawing.Size(242, 21)
        Me.cb_Request_Reason.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(15, 116)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(134, 21)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "사유"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_lotNo
        '
        Me.tb_lotNo.BackColor = System.Drawing.SystemColors.Window
        Me.tb_lotNo.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_lotNo.Location = New System.Drawing.Point(149, 24)
        Me.tb_lotNo.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_lotNo.Name = "tb_lotNo"
        Me.tb_lotNo.Size = New System.Drawing.Size(242, 21)
        Me.tb_lotNo.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(15, 24)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(134, 21)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Lot No. or YJ No."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cb_Process
        '
        Me.cb_Process.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_Process.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.cb_Process.FormattingEnabled = True
        Me.cb_Process.Items.AddRange(New Object() {"제전박스(수입검사)", "Baking JIG", "Tray Box(Bake)", "Carrier JIG", "Tray Box(외관불량)", "Tray Box(기능불량)", "폐기자재"})
        Me.cb_Process.Location = New System.Drawing.Point(149, 70)
        Me.cb_Process.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.cb_Process.Name = "cb_Process"
        Me.cb_Process.Size = New System.Drawing.Size(202, 21)
        Me.cb_Process.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(15, 70)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(134, 21)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "공정"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.SplitContainer1.Panel1.Controls.Add(Me.btn_TestLabel)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label32)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label18)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 936)
        Me.SplitContainer1.SplitterDistance = 421
        Me.SplitContainer1.TabIndex = 1
        '
        'btn_TestLabel
        '
        Me.btn_TestLabel.BackColor = System.Drawing.Color.Red
        Me.btn_TestLabel.Font = New System.Drawing.Font("굴림", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btn_TestLabel.Image = Global.Repair_System.My.Resources.Resources.checkmark
        Me.btn_TestLabel.Location = New System.Drawing.Point(156, 398)
        Me.btn_TestLabel.Name = "btn_TestLabel"
        Me.btn_TestLabel.Size = New System.Drawing.Size(207, 87)
        Me.btn_TestLabel.TabIndex = 3
        Me.btn_TestLabel.Text = "테스트 라벨 발행"
        Me.btn_TestLabel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btn_TestLabel.UseVisualStyleBackColor = True
        '
        'Label32
        '
        Me.Label32.BackColor = System.Drawing.Color.White
        Me.Label32.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label32.Font = New System.Drawing.Font("굴림", 16.0!, System.Drawing.FontStyle.Bold)
        Me.Label32.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label32.Location = New System.Drawing.Point(27, 26)
        Me.Label32.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(345, 47)
        Me.Label32.TabIndex = 0
        Me.Label32.Text = "프린터 설정"
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel5
        '
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel5.Controls.Add(Me.cb_PrinterList)
        Me.Panel5.Controls.Add(Me.Label23)
        Me.Panel5.Controls.Add(Me.CB_Print_Use)
        Me.Panel5.Controls.Add(Me.TB_StopBits)
        Me.Panel5.Controls.Add(Me.Label34)
        Me.Panel5.Controls.Add(Me.TB_Parity)
        Me.Panel5.Controls.Add(Me.Label33)
        Me.Panel5.Controls.Add(Me.cb_Cable)
        Me.Panel5.Controls.Add(Me.TB_DataBits)
        Me.Panel5.Controls.Add(Me.Label26)
        Me.Panel5.Controls.Add(Me.Label27)
        Me.Panel5.Controls.Add(Me.Label28)
        Me.Panel5.Controls.Add(Me.TB_TOP_Loc)
        Me.Panel5.Controls.Add(Me.Label29)
        Me.Panel5.Controls.Add(Me.TB_BaudRate)
        Me.Panel5.Controls.Add(Me.TB_LEFT_Loc)
        Me.Panel5.Controls.Add(Me.Label30)
        Me.Panel5.Controls.Add(Me.Label31)
        Me.Panel5.Controls.Add(Me.TB_Port)
        Me.Panel5.Location = New System.Drawing.Point(27, 77)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(345, 315)
        Me.Panel5.TabIndex = 2
        '
        'cb_PrinterList
        '
        Me.cb_PrinterList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_PrinterList.Font = New System.Drawing.Font("굴림", 13.75!, System.Drawing.FontStyle.Bold)
        Me.cb_PrinterList.FormattingEnabled = True
        Me.cb_PrinterList.Location = New System.Drawing.Point(180, 78)
        Me.cb_PrinterList.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.cb_PrinterList.Name = "cb_PrinterList"
        Me.cb_PrinterList.Size = New System.Drawing.Size(154, 26)
        Me.cb_PrinterList.TabIndex = 4
        '
        'Label23
        '
        Me.Label23.BackColor = System.Drawing.Color.SlateGray
        Me.Label23.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label23.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label23.ForeColor = System.Drawing.Color.White
        Me.Label23.Location = New System.Drawing.Point(9, 78)
        Me.Label23.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(171, 26)
        Me.Label23.TabIndex = 3
        Me.Label23.Text = "Printer 목록"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CB_Print_Use
        '
        Me.CB_Print_Use.AutoSize = True
        Me.CB_Print_Use.Checked = True
        Me.CB_Print_Use.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CB_Print_Use.Location = New System.Drawing.Point(9, 30)
        Me.CB_Print_Use.Name = "CB_Print_Use"
        Me.CB_Print_Use.Size = New System.Drawing.Size(76, 16)
        Me.CB_Print_Use.TabIndex = 0
        Me.CB_Print_Use.Text = "라벨 발행"
        Me.CB_Print_Use.UseVisualStyleBackColor = True
        '
        'TB_StopBits
        '
        Me.TB_StopBits.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_StopBits.Location = New System.Drawing.Point(180, 274)
        Me.TB_StopBits.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_StopBits.Name = "TB_StopBits"
        Me.TB_StopBits.Size = New System.Drawing.Size(154, 26)
        Me.TB_StopBits.TabIndex = 18
        Me.TB_StopBits.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label34
        '
        Me.Label34.BackColor = System.Drawing.Color.SlateGray
        Me.Label34.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label34.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label34.ForeColor = System.Drawing.Color.White
        Me.Label34.Location = New System.Drawing.Point(9, 274)
        Me.Label34.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(171, 26)
        Me.Label34.TabIndex = 17
        Me.Label34.Text = "StopBits"
        Me.Label34.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Parity
        '
        Me.TB_Parity.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_Parity.Location = New System.Drawing.Point(180, 246)
        Me.TB_Parity.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Parity.Name = "TB_Parity"
        Me.TB_Parity.Size = New System.Drawing.Size(154, 26)
        Me.TB_Parity.TabIndex = 16
        Me.TB_Parity.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label33
        '
        Me.Label33.BackColor = System.Drawing.Color.SlateGray
        Me.Label33.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label33.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label33.ForeColor = System.Drawing.Color.White
        Me.Label33.Location = New System.Drawing.Point(9, 246)
        Me.Label33.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(171, 26)
        Me.Label33.TabIndex = 15
        Me.Label33.Text = "Parity"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cb_Cable
        '
        Me.cb_Cable.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_Cable.Font = New System.Drawing.Font("굴림", 13.75!, System.Drawing.FontStyle.Bold)
        Me.cb_Cable.FormattingEnabled = True
        Me.cb_Cable.Items.AddRange(New Object() {"COM", "LPT", "USB"})
        Me.cb_Cable.Location = New System.Drawing.Point(180, 50)
        Me.cb_Cable.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.cb_Cable.Name = "cb_Cable"
        Me.cb_Cable.Size = New System.Drawing.Size(154, 26)
        Me.cb_Cable.TabIndex = 2
        '
        'TB_DataBits
        '
        Me.TB_DataBits.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_DataBits.Location = New System.Drawing.Point(180, 218)
        Me.TB_DataBits.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_DataBits.Name = "TB_DataBits"
        Me.TB_DataBits.Size = New System.Drawing.Size(154, 26)
        Me.TB_DataBits.TabIndex = 14
        Me.TB_DataBits.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.Color.SlateGray
        Me.Label26.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label26.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.White
        Me.Label26.Location = New System.Drawing.Point(9, 50)
        Me.Label26.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(171, 26)
        Me.Label26.TabIndex = 1
        Me.Label26.Text = "COM / LPT / USB"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label27
        '
        Me.Label27.BackColor = System.Drawing.Color.SlateGray
        Me.Label27.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label27.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label27.ForeColor = System.Drawing.Color.White
        Me.Label27.Location = New System.Drawing.Point(9, 218)
        Me.Label27.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(171, 26)
        Me.Label27.TabIndex = 13
        Me.Label27.Text = "DataBits"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.Color.SlateGray
        Me.Label28.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label28.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label28.ForeColor = System.Drawing.Color.White
        Me.Label28.Location = New System.Drawing.Point(9, 162)
        Me.Label28.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(171, 26)
        Me.Label28.TabIndex = 9
        Me.Label28.Text = "인쇄위치(세로)"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_TOP_Loc
        '
        Me.TB_TOP_Loc.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_TOP_Loc.Location = New System.Drawing.Point(180, 162)
        Me.TB_TOP_Loc.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_TOP_Loc.Name = "TB_TOP_Loc"
        Me.TB_TOP_Loc.Size = New System.Drawing.Size(154, 26)
        Me.TB_TOP_Loc.TabIndex = 10
        Me.TB_TOP_Loc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label29
        '
        Me.Label29.BackColor = System.Drawing.Color.SlateGray
        Me.Label29.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label29.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label29.ForeColor = System.Drawing.Color.White
        Me.Label29.Location = New System.Drawing.Point(9, 190)
        Me.Label29.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(171, 26)
        Me.Label29.TabIndex = 11
        Me.Label29.Text = "BaudRate"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_BaudRate
        '
        Me.TB_BaudRate.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_BaudRate.Location = New System.Drawing.Point(180, 190)
        Me.TB_BaudRate.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_BaudRate.Name = "TB_BaudRate"
        Me.TB_BaudRate.Size = New System.Drawing.Size(154, 26)
        Me.TB_BaudRate.TabIndex = 12
        Me.TB_BaudRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TB_LEFT_Loc
        '
        Me.TB_LEFT_Loc.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_LEFT_Loc.Location = New System.Drawing.Point(180, 134)
        Me.TB_LEFT_Loc.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_LEFT_Loc.Name = "TB_LEFT_Loc"
        Me.TB_LEFT_Loc.Size = New System.Drawing.Size(154, 26)
        Me.TB_LEFT_Loc.TabIndex = 8
        Me.TB_LEFT_Loc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label30
        '
        Me.Label30.BackColor = System.Drawing.Color.SlateGray
        Me.Label30.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label30.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label30.ForeColor = System.Drawing.Color.White
        Me.Label30.Location = New System.Drawing.Point(9, 106)
        Me.Label30.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(171, 26)
        Me.Label30.TabIndex = 5
        Me.Label30.Text = "Port"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label31
        '
        Me.Label31.BackColor = System.Drawing.Color.SlateGray
        Me.Label31.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label31.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label31.ForeColor = System.Drawing.Color.White
        Me.Label31.Location = New System.Drawing.Point(9, 134)
        Me.Label31.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(171, 26)
        Me.Label31.TabIndex = 7
        Me.Label31.Text = "인쇄위치(가로)"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Port
        '
        Me.TB_Port.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_Port.Location = New System.Drawing.Point(180, 106)
        Me.TB_Port.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Port.Name = "TB_Port"
        Me.TB_Port.Size = New System.Drawing.Size(154, 26)
        Me.TB_Port.TabIndex = 6
        Me.TB_Port.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(223, 915)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(186, 12)
        Me.Label18.TabIndex = 4
        Me.Label18.Text = "QR Code : YJ No / Lot No / Qty"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.grid_History)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label17)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.AutoScroll = True
        Me.SplitContainer2.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer2.Size = New System.Drawing.Size(839, 936)
        Me.SplitContainer2.SplitterDistance = 545
        Me.SplitContainer2.TabIndex = 3
        '
        'grid_History
        '
        Me.grid_History.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_History.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_History.Location = New System.Drawing.Point(0, 53)
        Me.grid_History.Name = "grid_History"
        Me.grid_History.Rows.Count = 2
        Me.grid_History.Rows.DefaultSize = 20
        Me.grid_History.Size = New System.Drawing.Size(839, 492)
        Me.grid_History.StyleInfo = resources.GetString("grid_History.StyleInfo")
        Me.grid_History.TabIndex = 1
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label17.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label17.Font = New System.Drawing.Font("굴림", 16.0!, System.Drawing.FontStyle.Bold)
        Me.Label17.ForeColor = System.Drawing.Color.White
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(839, 53)
        Me.Label17.TabIndex = 0
        Me.Label17.Text = "재발행 이력"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(839, 387)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(831, 361)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "공정 라벨"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.White
        Me.GroupBox2.Controls.Add(Me.tb_Lot_Option)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.tb_total_print_qty)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.tb_After_RCD_Marking)
        Me.GroupBox2.Controls.Add(Me.tb_Before_RCD_Marking)
        Me.GroupBox2.Controls.Add(Me.tb_After_PMIC_Marking)
        Me.GroupBox2.Controls.Add(Me.tb_Before_PMIC_Marking)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.tb_After_RCD)
        Me.GroupBox2.Controls.Add(Me.tb_Before_RCD)
        Me.GroupBox2.Controls.Add(Me.tb_After_PMIC)
        Me.GroupBox2.Controls.Add(Me.tb_Before_PMIC)
        Me.GroupBox2.Controls.Add(Me.Label37)
        Me.GroupBox2.Controls.Add(Me.tb_After_Bucket)
        Me.GroupBox2.Controls.Add(Me.tb_before_Bucket)
        Me.GroupBox2.Controls.Add(Me.Label36)
        Me.GroupBox2.Controls.Add(Me.Label35)
        Me.GroupBox2.Controls.Add(Me.Label25)
        Me.GroupBox2.Controls.Add(Me.tb_Print_YJ_No)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.tb_Print_Lot_No)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(6, 159)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(824, 147)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "발행 내용"
        '
        'tb_Lot_Option
        '
        Me.tb_Lot_Option.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Lot_Option.Enabled = False
        Me.tb_Lot_Option.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Lot_Option.Location = New System.Drawing.Point(551, 46)
        Me.tb_Lot_Option.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Lot_Option.Name = "tb_Lot_Option"
        Me.tb_Lot_Option.Size = New System.Drawing.Size(242, 21)
        Me.tb_Lot_Option.TabIndex = 22
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label10.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(417, 46)
        Me.Label10.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(134, 21)
        Me.Label10.TabIndex = 21
        Me.Label10.Text = "Option"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_total_print_qty
        '
        Me.tb_total_print_qty.BackColor = System.Drawing.SystemColors.Window
        Me.tb_total_print_qty.Enabled = False
        Me.tb_total_print_qty.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_total_print_qty.Location = New System.Drawing.Point(551, 23)
        Me.tb_total_print_qty.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_total_print_qty.Name = "tb_total_print_qty"
        Me.tb_total_print_qty.Size = New System.Drawing.Size(242, 21)
        Me.tb_total_print_qty.TabIndex = 20
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(417, 23)
        Me.Label7.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(134, 21)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "총 인쇄 수량"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_After_RCD_Marking
        '
        Me.tb_After_RCD_Marking.BackColor = System.Drawing.SystemColors.Window
        Me.tb_After_RCD_Marking.Enabled = False
        Me.tb_After_RCD_Marking.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_After_RCD_Marking.Location = New System.Drawing.Point(685, 118)
        Me.tb_After_RCD_Marking.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_After_RCD_Marking.Name = "tb_After_RCD_Marking"
        Me.tb_After_RCD_Marking.Size = New System.Drawing.Size(134, 21)
        Me.tb_After_RCD_Marking.TabIndex = 18
        Me.tb_After_RCD_Marking.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_Before_RCD_Marking
        '
        Me.tb_Before_RCD_Marking.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Before_RCD_Marking.Enabled = False
        Me.tb_Before_RCD_Marking.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Before_RCD_Marking.Location = New System.Drawing.Point(685, 95)
        Me.tb_Before_RCD_Marking.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Before_RCD_Marking.Name = "tb_Before_RCD_Marking"
        Me.tb_Before_RCD_Marking.Size = New System.Drawing.Size(134, 21)
        Me.tb_Before_RCD_Marking.TabIndex = 12
        Me.tb_Before_RCD_Marking.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_After_PMIC_Marking
        '
        Me.tb_After_PMIC_Marking.BackColor = System.Drawing.SystemColors.Window
        Me.tb_After_PMIC_Marking.Enabled = False
        Me.tb_After_PMIC_Marking.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_After_PMIC_Marking.Location = New System.Drawing.Point(417, 118)
        Me.tb_After_PMIC_Marking.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_After_PMIC_Marking.Name = "tb_After_PMIC_Marking"
        Me.tb_After_PMIC_Marking.Size = New System.Drawing.Size(134, 21)
        Me.tb_After_PMIC_Marking.TabIndex = 16
        Me.tb_After_PMIC_Marking.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_Before_PMIC_Marking
        '
        Me.tb_Before_PMIC_Marking.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Before_PMIC_Marking.Enabled = False
        Me.tb_Before_PMIC_Marking.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Before_PMIC_Marking.Location = New System.Drawing.Point(417, 95)
        Me.tb_Before_PMIC_Marking.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Before_PMIC_Marking.Name = "tb_Before_PMIC_Marking"
        Me.tb_Before_PMIC_Marking.Size = New System.Drawing.Size(134, 21)
        Me.tb_Before_PMIC_Marking.TabIndex = 10
        Me.tb_Before_PMIC_Marking.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(551, 72)
        Me.Label6.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(268, 21)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "RCD"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_After_RCD
        '
        Me.tb_After_RCD.BackColor = System.Drawing.SystemColors.Window
        Me.tb_After_RCD.Enabled = False
        Me.tb_After_RCD.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_After_RCD.Location = New System.Drawing.Point(551, 118)
        Me.tb_After_RCD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_After_RCD.Name = "tb_After_RCD"
        Me.tb_After_RCD.Size = New System.Drawing.Size(134, 21)
        Me.tb_After_RCD.TabIndex = 17
        Me.tb_After_RCD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_Before_RCD
        '
        Me.tb_Before_RCD.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Before_RCD.Enabled = False
        Me.tb_Before_RCD.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Before_RCD.Location = New System.Drawing.Point(551, 95)
        Me.tb_Before_RCD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Before_RCD.Name = "tb_Before_RCD"
        Me.tb_Before_RCD.Size = New System.Drawing.Size(134, 21)
        Me.tb_Before_RCD.TabIndex = 11
        Me.tb_Before_RCD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_After_PMIC
        '
        Me.tb_After_PMIC.BackColor = System.Drawing.SystemColors.Window
        Me.tb_After_PMIC.Enabled = False
        Me.tb_After_PMIC.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_After_PMIC.Location = New System.Drawing.Point(283, 118)
        Me.tb_After_PMIC.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_After_PMIC.Name = "tb_After_PMIC"
        Me.tb_After_PMIC.Size = New System.Drawing.Size(134, 21)
        Me.tb_After_PMIC.TabIndex = 15
        Me.tb_After_PMIC.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_Before_PMIC
        '
        Me.tb_Before_PMIC.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Before_PMIC.Enabled = False
        Me.tb_Before_PMIC.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Before_PMIC.Location = New System.Drawing.Point(283, 95)
        Me.tb_Before_PMIC.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Before_PMIC.Name = "tb_Before_PMIC"
        Me.tb_Before_PMIC.Size = New System.Drawing.Size(134, 21)
        Me.tb_Before_PMIC.TabIndex = 9
        Me.tb_Before_PMIC.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label37
        '
        Me.Label37.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label37.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label37.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label37.ForeColor = System.Drawing.Color.White
        Me.Label37.Location = New System.Drawing.Point(283, 72)
        Me.Label37.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(268, 21)
        Me.Label37.TabIndex = 5
        Me.Label37.Text = "PMIC"
        Me.Label37.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_After_Bucket
        '
        Me.tb_After_Bucket.BackColor = System.Drawing.SystemColors.Window
        Me.tb_After_Bucket.Enabled = False
        Me.tb_After_Bucket.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_After_Bucket.Location = New System.Drawing.Point(149, 118)
        Me.tb_After_Bucket.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_After_Bucket.Name = "tb_After_Bucket"
        Me.tb_After_Bucket.Size = New System.Drawing.Size(134, 21)
        Me.tb_After_Bucket.TabIndex = 14
        Me.tb_After_Bucket.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_before_Bucket
        '
        Me.tb_before_Bucket.BackColor = System.Drawing.SystemColors.Window
        Me.tb_before_Bucket.Enabled = False
        Me.tb_before_Bucket.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_before_Bucket.Location = New System.Drawing.Point(149, 95)
        Me.tb_before_Bucket.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_before_Bucket.Name = "tb_before_Bucket"
        Me.tb_before_Bucket.Size = New System.Drawing.Size(134, 21)
        Me.tb_before_Bucket.TabIndex = 8
        Me.tb_before_Bucket.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label36
        '
        Me.Label36.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label36.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label36.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label36.ForeColor = System.Drawing.Color.White
        Me.Label36.Location = New System.Drawing.Point(149, 72)
        Me.Label36.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(134, 21)
        Me.Label36.TabIndex = 4
        Me.Label36.Text = "Customer Code"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label35
        '
        Me.Label35.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label35.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label35.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label35.ForeColor = System.Drawing.Color.White
        Me.Label35.Location = New System.Drawing.Point(15, 118)
        Me.Label35.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(134, 21)
        Me.Label35.TabIndex = 13
        Me.Label35.Text = "변경"
        Me.Label35.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label25.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label25.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.White
        Me.Label25.Location = New System.Drawing.Point(15, 95)
        Me.Label25.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(134, 21)
        Me.Label25.TabIndex = 7
        Me.Label25.Text = "기존"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_Print_YJ_No
        '
        Me.tb_Print_YJ_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Print_YJ_No.Enabled = False
        Me.tb_Print_YJ_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Print_YJ_No.Location = New System.Drawing.Point(149, 46)
        Me.tb_Print_YJ_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Print_YJ_No.Name = "tb_Print_YJ_No"
        Me.tb_Print_YJ_No.Size = New System.Drawing.Size(242, 21)
        Me.tb_Print_YJ_No.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(15, 46)
        Me.Label5.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(134, 21)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "YJ No."
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Print_Lot_No
        '
        Me.tb_Print_Lot_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Print_Lot_No.Enabled = False
        Me.tb_Print_Lot_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Print_Lot_No.Location = New System.Drawing.Point(149, 23)
        Me.tb_Print_Lot_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Print_Lot_No.Name = "tb_Print_Lot_No"
        Me.tb_Print_Lot_No.Size = New System.Drawing.Size(242, 21)
        Me.tb_Print_Lot_No.TabIndex = 1
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(15, 23)
        Me.Label9.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(134, 21)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Lot No."
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox3)
        Me.TabPage2.Controls.Add(Me.GroupBox4)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(831, 361)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "자재 라벨"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.White
        Me.GroupBox3.Controls.Add(Me.btn_Parts_Save)
        Me.GroupBox3.Controls.Add(Me.tb_Parts_Requester)
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Controls.Add(Me.cb_Parts_Reason)
        Me.GroupBox3.Controls.Add(Me.Label13)
        Me.GroupBox3.Controls.Add(Me.tb_Parts_LotNo)
        Me.GroupBox3.Controls.Add(Me.Label14)
        Me.GroupBox3.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(6, 6)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(717, 103)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "발행 정보"
        '
        'btn_Parts_Save
        '
        Me.btn_Parts_Save.BackColor = System.Drawing.Color.Red
        Me.btn_Parts_Save.Font = New System.Drawing.Font("굴림", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btn_Parts_Save.Image = Global.Repair_System.My.Resources.Resources.barcode
        Me.btn_Parts_Save.Location = New System.Drawing.Point(551, 22)
        Me.btn_Parts_Save.Name = "btn_Parts_Save"
        Me.btn_Parts_Save.Size = New System.Drawing.Size(153, 69)
        Me.btn_Parts_Save.TabIndex = 1
        Me.btn_Parts_Save.Text = "저장 및 발행"
        Me.btn_Parts_Save.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btn_Parts_Save.UseVisualStyleBackColor = True
        '
        'tb_Parts_Requester
        '
        Me.tb_Parts_Requester.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Parts_Requester.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Parts_Requester.Location = New System.Drawing.Point(149, 47)
        Me.tb_Parts_Requester.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Parts_Requester.Name = "tb_Parts_Requester"
        Me.tb_Parts_Requester.Size = New System.Drawing.Size(242, 21)
        Me.tb_Parts_Requester.TabIndex = 3
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label12.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.White
        Me.Label12.Location = New System.Drawing.Point(15, 47)
        Me.Label12.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(134, 21)
        Me.Label12.TabIndex = 2
        Me.Label12.Text = "요청자"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cb_Parts_Reason
        '
        Me.cb_Parts_Reason.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.cb_Parts_Reason.FormattingEnabled = True
        Me.cb_Parts_Reason.Items.AddRange(New Object() {"라벨 훼손", "라벨 분실"})
        Me.cb_Parts_Reason.Location = New System.Drawing.Point(149, 70)
        Me.cb_Parts_Reason.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.cb_Parts_Reason.Name = "cb_Parts_Reason"
        Me.cb_Parts_Reason.Size = New System.Drawing.Size(242, 21)
        Me.cb_Parts_Reason.TabIndex = 8
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label13.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.White
        Me.Label13.Location = New System.Drawing.Point(15, 70)
        Me.Label13.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(134, 21)
        Me.Label13.TabIndex = 7
        Me.Label13.Text = "사유"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Parts_LotNo
        '
        Me.tb_Parts_LotNo.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Parts_LotNo.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Parts_LotNo.Location = New System.Drawing.Point(149, 24)
        Me.tb_Parts_LotNo.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Parts_LotNo.Name = "tb_Parts_LotNo"
        Me.tb_Parts_LotNo.Size = New System.Drawing.Size(242, 21)
        Me.tb_Parts_LotNo.TabIndex = 1
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label14.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(15, 24)
        Me.Label14.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(134, 21)
        Me.Label14.TabIndex = 0
        Me.Label14.Text = "Lot No. or YJ No."
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.Color.White
        Me.GroupBox4.Controls.Add(Me.TextBox1)
        Me.GroupBox4.Controls.Add(Me.Label20)
        Me.GroupBox4.Controls.Add(Me.tb_Parts_Load_TopMarking)
        Me.GroupBox4.Controls.Add(Me.Label19)
        Me.GroupBox4.Controls.Add(Me.tb_Parts_Load_YJNo)
        Me.GroupBox4.Controls.Add(Me.Label16)
        Me.GroupBox4.Controls.Add(Me.tb_Parts_Load_LotNo)
        Me.GroupBox4.Controls.Add(Me.Label15)
        Me.GroupBox4.Controls.Add(Me.tb_Parts_Load_PartNo)
        Me.GroupBox4.Controls.Add(Me.Label11)
        Me.GroupBox4.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(6, 115)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(717, 124)
        Me.GroupBox4.TabIndex = 4
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "발행 내용"
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.SystemColors.Window
        Me.TextBox1.Enabled = False
        Me.TextBox1.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TextBox1.Location = New System.Drawing.Point(525, 23)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(179, 21)
        Me.TextBox1.TabIndex = 11
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label20.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label20.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.White
        Me.Label20.Location = New System.Drawing.Point(391, 23)
        Me.Label20.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(134, 21)
        Me.Label20.TabIndex = 10
        Me.Label20.Text = "구분"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Parts_Load_TopMarking
        '
        Me.tb_Parts_Load_TopMarking.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Parts_Load_TopMarking.Enabled = False
        Me.tb_Parts_Load_TopMarking.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Parts_Load_TopMarking.Location = New System.Drawing.Point(149, 92)
        Me.tb_Parts_Load_TopMarking.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Parts_Load_TopMarking.Name = "tb_Parts_Load_TopMarking"
        Me.tb_Parts_Load_TopMarking.Size = New System.Drawing.Size(242, 21)
        Me.tb_Parts_Load_TopMarking.TabIndex = 9
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label19.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.White
        Me.Label19.Location = New System.Drawing.Point(15, 92)
        Me.Label19.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(134, 21)
        Me.Label19.TabIndex = 8
        Me.Label19.Text = "Top Marking"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Parts_Load_YJNo
        '
        Me.tb_Parts_Load_YJNo.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Parts_Load_YJNo.Enabled = False
        Me.tb_Parts_Load_YJNo.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Parts_Load_YJNo.Location = New System.Drawing.Point(149, 69)
        Me.tb_Parts_Load_YJNo.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Parts_Load_YJNo.Name = "tb_Parts_Load_YJNo"
        Me.tb_Parts_Load_YJNo.Size = New System.Drawing.Size(242, 21)
        Me.tb_Parts_Load_YJNo.TabIndex = 7
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label16.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.White
        Me.Label16.Location = New System.Drawing.Point(15, 69)
        Me.Label16.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(134, 21)
        Me.Label16.TabIndex = 6
        Me.Label16.Text = "YJ No."
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Parts_Load_LotNo
        '
        Me.tb_Parts_Load_LotNo.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Parts_Load_LotNo.Enabled = False
        Me.tb_Parts_Load_LotNo.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Parts_Load_LotNo.Location = New System.Drawing.Point(149, 46)
        Me.tb_Parts_Load_LotNo.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Parts_Load_LotNo.Name = "tb_Parts_Load_LotNo"
        Me.tb_Parts_Load_LotNo.Size = New System.Drawing.Size(242, 21)
        Me.tb_Parts_Load_LotNo.TabIndex = 5
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label15.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(15, 46)
        Me.Label15.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(134, 21)
        Me.Label15.TabIndex = 4
        Me.Label15.Text = "Lot No."
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Parts_Load_PartNo
        '
        Me.tb_Parts_Load_PartNo.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Parts_Load_PartNo.Enabled = False
        Me.tb_Parts_Load_PartNo.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Parts_Load_PartNo.Location = New System.Drawing.Point(149, 23)
        Me.tb_Parts_Load_PartNo.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Parts_Load_PartNo.Name = "tb_Parts_Load_PartNo"
        Me.tb_Parts_Load_PartNo.Size = New System.Drawing.Size(242, 21)
        Me.tb_Parts_Load_PartNo.TabIndex = 3
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label11.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(15, 23)
        Me.Label11.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(134, 21)
        Me.Label11.TabIndex = 2
        Me.Label11.Text = "Part No."
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Toolstrip1
        '
        Me.Toolstrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.Toolstrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Form_CLose, Me.BTN_Setting_Open})
        Me.Toolstrip1.Location = New System.Drawing.Point(0, 0)
        Me.Toolstrip1.Name = "Toolstrip1"
        Me.Toolstrip1.Size = New System.Drawing.Size(1264, 25)
        Me.Toolstrip1.TabIndex = 0
        '
        'Form_CLose
        '
        Me.Form_CLose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.Form_CLose.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.Form_CLose.Image = Global.Repair_System.My.Resources.Resources.Close
        Me.Form_CLose.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.Form_CLose.Name = "Form_CLose"
        Me.Form_CLose.Size = New System.Drawing.Size(23, 22)
        Me.Form_CLose.Text = "폼 닫기"
        '
        'BTN_Setting_Open
        '
        Me.BTN_Setting_Open.Image = Global.Repair_System.My.Resources.Resources.Settings
        Me.BTN_Setting_Open.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_Setting_Open.Name = "BTN_Setting_Open"
        Me.BTN_Setting_Open.Size = New System.Drawing.Size(119, 22)
        Me.BTN_Setting_Open.Text = "프린터 설정 열기"
        '
        'frm_IQC_RePrint
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1264, 961)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.Toolstrip1)
        Me.Name = "frm_IQC_RePrint"
        Me.Text = "라벨 추가 발행 및 재발행"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.grid_History, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.Toolstrip1.ResumeLayout(False)
        Me.Toolstrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cb_Process As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents grid_History As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents BTN_SavePrint As System.Windows.Forms.Button
    Friend WithEvents Toolstrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents Form_CLose As System.Windows.Forms.ToolStripButton
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents TB_DataBits As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents TB_TOP_Loc As System.Windows.Forms.TextBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents TB_BaudRate As System.Windows.Forms.TextBox
    Friend WithEvents TB_LEFT_Loc As System.Windows.Forms.TextBox
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents TB_Port As System.Windows.Forms.TextBox
    Friend WithEvents cb_Cable As System.Windows.Forms.ComboBox
    Friend WithEvents TB_StopBits As System.Windows.Forms.TextBox
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents TB_Parity As System.Windows.Forms.TextBox
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents BTN_Setting_Open As System.Windows.Forms.ToolStripButton
    Friend WithEvents CB_Print_Use As System.Windows.Forms.CheckBox
    Friend WithEvents cb_PrinterList As ComboBox
    Friend WithEvents Label23 As Label
    Friend WithEvents btn_TestLabel As Button
    Friend WithEvents cb_Request_Reason As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents tb_lotNo As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents tb_Requester As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents tb_Process_Char As TextBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents tb_Print_Lot_No As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents tb_Print_YJ_No As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents tb_After_RCD_Marking As TextBox
    Friend WithEvents tb_Before_RCD_Marking As TextBox
    Friend WithEvents tb_After_PMIC_Marking As TextBox
    Friend WithEvents tb_Before_PMIC_Marking As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents tb_After_RCD As TextBox
    Friend WithEvents tb_Before_RCD As TextBox
    Friend WithEvents tb_After_PMIC As TextBox
    Friend WithEvents tb_Before_PMIC As TextBox
    Friend WithEvents Label37 As Label
    Friend WithEvents tb_After_Bucket As TextBox
    Friend WithEvents tb_before_Bucket As TextBox
    Friend WithEvents Label36 As Label
    Friend WithEvents Label35 As Label
    Friend WithEvents Label25 As Label
    Friend WithEvents tb_total_print_qty As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents tb_Process_No As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents tb_Lot_Option As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents btn_Parts_Save As Button
    Friend WithEvents tb_Parts_Requester As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents cb_Parts_Reason As ComboBox
    Friend WithEvents Label13 As Label
    Friend WithEvents tb_Parts_LotNo As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents tb_Parts_Load_TopMarking As TextBox
    Friend WithEvents Label19 As Label
    Friend WithEvents tb_Parts_Load_YJNo As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents tb_Parts_Load_LotNo As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents tb_Parts_Load_PartNo As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label20 As Label
    Friend WithEvents Label21 As Label
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_Parts_Kitting
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Parts_Kitting))
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_Search = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_New_Kitting = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Save = New System.Windows.Forms.ToolStripButton()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
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
        Me.grid_KittingList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.tb_S_YJ_No = New System.Windows.Forms.TextBox()
        Me.tb_S_Lot_No = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.btn_Save2 = New System.Windows.Forms.Button()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.btn_Check_Print = New System.Windows.Forms.Button()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.tb_Parts_Qty = New System.Windows.Forms.TextBox()
        Me.tb_Parts_Barcode = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.tb_Parts_YJ_No = New System.Windows.Forms.TextBox()
        Me.tb_Parts_PartNo = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.tb_Parts_Lot_No = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.grid_PartsList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.tb_Option = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.tb_Lot_Status = New System.Windows.Forms.TextBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.tb_After_RCD = New System.Windows.Forms.TextBox()
        Me.tb_Before_RCD = New System.Windows.Forms.TextBox()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.tb_After_PMIC = New System.Windows.Forms.TextBox()
        Me.tb_Before_PMIC = New System.Windows.Forms.TextBox()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.tb_After_CC = New System.Windows.Forms.TextBox()
        Me.tb_before_CC = New System.Windows.Forms.TextBox()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.tb_ORG_Product = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tb_Lot_Qty = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tb_YJ_No = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.tb_Product = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tb_Lot_No = New System.Windows.Forms.TextBox()
        Me.tb_Barcode = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel5.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.grid_KittingList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.grid_PartsList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripSeparator3, Me.BTN_Search, Me.ToolStripSeparator1, Me.btn_New_Kitting, Me.ToolStripSeparator2, Me.btn_Save, Me.Form_CLose})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1484, 25)
        Me.ToolStrip1.TabIndex = 5
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.Image = Global.Repair_System.My.Resources.Resources.Settings
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(91, 22)
        Me.ToolStripButton1.Text = "프린터 설정"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
        '
        'BTN_Search
        '
        Me.BTN_Search.Image = Global.Repair_System.My.Resources.Resources.Search_121
        Me.BTN_Search.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_Search.Name = "BTN_Search"
        Me.BTN_Search.Size = New System.Drawing.Size(51, 22)
        Me.BTN_Search.Text = "조회"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'btn_New_Kitting
        '
        Me.btn_New_Kitting.Image = Global.Repair_System.My.Resources.Resources.TEST_FOLDER
        Me.btn_New_Kitting.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_New_Kitting.Name = "btn_New_Kitting"
        Me.btn_New_Kitting.Size = New System.Drawing.Size(75, 22)
        Me.btn_New_Kitting.Text = "신규등록"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'btn_Save
        '
        Me.btn_Save.Enabled = False
        Me.btn_Save.Image = Global.Repair_System.My.Resources.Resources.save2
        Me.btn_Save.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_Save.Name = "btn_Save"
        Me.btn_Save.Size = New System.Drawing.Size(51, 22)
        Me.btn_Save.Text = "저장"
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
        Me.SplitContainer1.Panel1Collapsed = True
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1484, 748)
        Me.SplitContainer1.SplitterDistance = 421
        Me.SplitContainer1.TabIndex = 6
        '
        'btn_TestLabel
        '
        Me.btn_TestLabel.BackColor = System.Drawing.Color.Red
        Me.btn_TestLabel.Font = New System.Drawing.Font("굴림", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btn_TestLabel.Image = Global.Repair_System.My.Resources.Resources.checkmark
        Me.btn_TestLabel.Location = New System.Drawing.Point(160, 384)
        Me.btn_TestLabel.Name = "btn_TestLabel"
        Me.btn_TestLabel.Size = New System.Drawing.Size(207, 87)
        Me.btn_TestLabel.TabIndex = 18
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
        Me.Label32.Location = New System.Drawing.Point(31, 12)
        Me.Label32.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(345, 47)
        Me.Label32.TabIndex = 17
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
        Me.Panel5.Location = New System.Drawing.Point(31, 63)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(345, 315)
        Me.Panel5.TabIndex = 16
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
        Me.cb_PrinterList.TabIndex = 22
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
        Me.Label23.TabIndex = 21
        Me.Label23.Text = "Printer 목록"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CB_Print_Use
        '
        Me.CB_Print_Use.AutoSize = True
        Me.CB_Print_Use.Location = New System.Drawing.Point(9, 30)
        Me.CB_Print_Use.Name = "CB_Print_Use"
        Me.CB_Print_Use.Size = New System.Drawing.Size(76, 16)
        Me.CB_Print_Use.TabIndex = 20
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
        Me.TB_StopBits.TabIndex = 19
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
        Me.Label34.TabIndex = 18
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
        Me.TB_Parity.TabIndex = 17
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
        Me.Label33.TabIndex = 16
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
        Me.cb_Cable.TabIndex = 1
        '
        'TB_DataBits
        '
        Me.TB_DataBits.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_DataBits.Location = New System.Drawing.Point(180, 218)
        Me.TB_DataBits.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_DataBits.Name = "TB_DataBits"
        Me.TB_DataBits.Size = New System.Drawing.Size(154, 26)
        Me.TB_DataBits.TabIndex = 15
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
        Me.Label26.TabIndex = 0
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
        Me.Label27.TabIndex = 14
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
        Me.Label28.TabIndex = 8
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
        Me.TB_TOP_Loc.TabIndex = 9
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
        Me.TB_LEFT_Loc.TabIndex = 6
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
        Me.Label30.TabIndex = 2
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
        Me.Label31.TabIndex = 5
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
        Me.TB_Port.TabIndex = 3
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
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.grid_KittingList)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Panel1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.BackColor = System.Drawing.Color.White
        Me.SplitContainer2.Panel2.Controls.Add(Me.btn_Save2)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer2.Panel2.Controls.Add(Me.grid_PartsList)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer2.Size = New System.Drawing.Size(1484, 748)
        Me.SplitContainer2.SplitterDistance = 354
        Me.SplitContainer2.TabIndex = 1
        '
        'grid_KittingList
        '
        Me.grid_KittingList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_KittingList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_KittingList.Location = New System.Drawing.Point(0, 85)
        Me.grid_KittingList.Name = "grid_KittingList"
        Me.grid_KittingList.Rows.Count = 2
        Me.grid_KittingList.Rows.DefaultSize = 20
        Me.grid_KittingList.Size = New System.Drawing.Size(354, 663)
        Me.grid_KittingList.StyleInfo = resources.GetString("grid_KittingList.StyleInfo")
        Me.grid_KittingList.TabIndex = 2
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Panel1.Controls.Add(Me.tb_S_YJ_No)
        Me.Panel1.Controls.Add(Me.tb_S_Lot_No)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.DateTimePicker2)
        Me.Panel1.Controls.Add(Me.DateTimePicker1)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(354, 85)
        Me.Panel1.TabIndex = 3
        '
        'tb_S_YJ_No
        '
        Me.tb_S_YJ_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_S_YJ_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_S_YJ_No.Location = New System.Drawing.Point(87, 31)
        Me.tb_S_YJ_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_S_YJ_No.Name = "tb_S_YJ_No"
        Me.tb_S_YJ_No.Size = New System.Drawing.Size(242, 21)
        Me.tb_S_YJ_No.TabIndex = 8
        '
        'tb_S_Lot_No
        '
        Me.tb_S_Lot_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_S_Lot_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_S_Lot_No.Location = New System.Drawing.Point(87, 54)
        Me.tb_S_Lot_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_S_Lot_No.Name = "tb_S_Lot_No"
        Me.tb_S_Lot_No.Size = New System.Drawing.Size(242, 21)
        Me.tb_S_Lot_No.TabIndex = 7
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(9, 54)
        Me.Label5.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(78, 21)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Lot No."
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(9, 31)
        Me.Label8.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(78, 21)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "YJ No."
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(193, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(30, 21)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "~"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker2.Location = New System.Drawing.Point(226, 8)
        Me.DateTimePicker2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(103, 21)
        Me.DateTimePicker2.TabIndex = 3
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(87, 8)
        Me.DateTimePicker1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(103, 21)
        Me.DateTimePicker1.TabIndex = 1
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label11.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(9, 8)
        Me.Label11.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(78, 21)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "Kitting 일자"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_Save2
        '
        Me.btn_Save2.Enabled = False
        Me.btn_Save2.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btn_Save2.Image = Global.Repair_System.My.Resources.Resources.save_5
        Me.btn_Save2.Location = New System.Drawing.Point(787, 550)
        Me.btn_Save2.Name = "btn_Save2"
        Me.btn_Save2.Size = New System.Drawing.Size(197, 91)
        Me.btn_Save2.TabIndex = 27
        Me.btn_Save2.Text = "Kitting 정보 저장"
        Me.btn_Save2.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btn_Save2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btn_Save2.UseVisualStyleBackColor = True
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.Label24)
        Me.Panel4.Controls.Add(Me.btn_Check_Print)
        Me.Panel4.Controls.Add(Me.Label21)
        Me.Panel4.Controls.Add(Me.Label15)
        Me.Panel4.Controls.Add(Me.tb_Parts_Qty)
        Me.Panel4.Controls.Add(Me.tb_Parts_Barcode)
        Me.Panel4.Controls.Add(Me.Label20)
        Me.Panel4.Controls.Add(Me.Label16)
        Me.Panel4.Controls.Add(Me.tb_Parts_YJ_No)
        Me.Panel4.Controls.Add(Me.tb_Parts_PartNo)
        Me.Panel4.Controls.Add(Me.Label19)
        Me.Panel4.Controls.Add(Me.Label17)
        Me.Panel4.Controls.Add(Me.tb_Parts_Lot_No)
        Me.Panel4.Location = New System.Drawing.Point(595, 337)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(392, 207)
        Me.Panel4.TabIndex = 26
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.DarkRed
        Me.Label24.Location = New System.Drawing.Point(5, 178)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(261, 12)
        Me.Label24.TabIndex = 27
        Me.Label24.Text = "※ 출력된 라벨은 컷팅된 자재에 붙혀 주십시오."
        '
        'btn_Check_Print
        '
        Me.btn_Check_Print.Location = New System.Drawing.Point(268, 168)
        Me.btn_Check_Print.Name = "btn_Check_Print"
        Me.btn_Check_Print.Size = New System.Drawing.Size(117, 32)
        Me.btn_Check_Print.TabIndex = 26
        Me.btn_Check_Print.Text = "확인 및 라벨 출력"
        Me.btn_Check_Print.UseVisualStyleBackColor = True
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold)
        Me.Label21.ForeColor = System.Drawing.Color.Black
        Me.Label21.Location = New System.Drawing.Point(3, 10)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(100, 24)
        Me.Label21.TabIndex = 25
        Me.Label21.Text = "[ 자재확인 ]"
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label15.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(7, 40)
        Me.Label15.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(134, 21)
        Me.Label15.TabIndex = 21
        Me.Label15.Text = "자재 Barcode"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Parts_Qty
        '
        Me.tb_Parts_Qty.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Parts_Qty.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Parts_Qty.Location = New System.Drawing.Point(141, 132)
        Me.tb_Parts_Qty.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Parts_Qty.Name = "tb_Parts_Qty"
        Me.tb_Parts_Qty.Size = New System.Drawing.Size(242, 21)
        Me.tb_Parts_Qty.TabIndex = 24
        '
        'tb_Parts_Barcode
        '
        Me.tb_Parts_Barcode.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Parts_Barcode.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Parts_Barcode.Location = New System.Drawing.Point(141, 40)
        Me.tb_Parts_Barcode.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Parts_Barcode.Name = "tb_Parts_Barcode"
        Me.tb_Parts_Barcode.Size = New System.Drawing.Size(242, 21)
        Me.tb_Parts_Barcode.TabIndex = 22
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label20.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label20.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.White
        Me.Label20.Location = New System.Drawing.Point(7, 132)
        Me.Label20.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(134, 21)
        Me.Label20.TabIndex = 23
        Me.Label20.Text = "사용수량"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label16.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.White
        Me.Label16.Location = New System.Drawing.Point(7, 63)
        Me.Label16.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(134, 21)
        Me.Label16.TabIndex = 21
        Me.Label16.Text = "Part No."
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Parts_YJ_No
        '
        Me.tb_Parts_YJ_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Parts_YJ_No.Enabled = False
        Me.tb_Parts_YJ_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Parts_YJ_No.Location = New System.Drawing.Point(141, 109)
        Me.tb_Parts_YJ_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Parts_YJ_No.Name = "tb_Parts_YJ_No"
        Me.tb_Parts_YJ_No.Size = New System.Drawing.Size(242, 21)
        Me.tb_Parts_YJ_No.TabIndex = 24
        '
        'tb_Parts_PartNo
        '
        Me.tb_Parts_PartNo.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Parts_PartNo.Enabled = False
        Me.tb_Parts_PartNo.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Parts_PartNo.Location = New System.Drawing.Point(141, 63)
        Me.tb_Parts_PartNo.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Parts_PartNo.Name = "tb_Parts_PartNo"
        Me.tb_Parts_PartNo.Size = New System.Drawing.Size(242, 21)
        Me.tb_Parts_PartNo.TabIndex = 22
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label19.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.White
        Me.Label19.Location = New System.Drawing.Point(7, 109)
        Me.Label19.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(134, 21)
        Me.Label19.TabIndex = 23
        Me.Label19.Text = "YJ No."
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label17.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.White
        Me.Label17.Location = New System.Drawing.Point(7, 86)
        Me.Label17.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(134, 21)
        Me.Label17.TabIndex = 23
        Me.Label17.Text = "Lot No."
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Parts_Lot_No
        '
        Me.tb_Parts_Lot_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Parts_Lot_No.Enabled = False
        Me.tb_Parts_Lot_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Parts_Lot_No.Location = New System.Drawing.Point(141, 86)
        Me.tb_Parts_Lot_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Parts_Lot_No.Name = "tb_Parts_Lot_No"
        Me.tb_Parts_Lot_No.Size = New System.Drawing.Size(242, 21)
        Me.tb_Parts_Lot_No.TabIndex = 24
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold)
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(15, 310)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(143, 24)
        Me.Label12.TabIndex = 20
        Me.Label12.Text = "[ 출고자재 List ]"
        '
        'grid_PartsList
        '
        Me.grid_PartsList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grid_PartsList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_PartsList.Location = New System.Drawing.Point(10, 337)
        Me.grid_PartsList.Name = "grid_PartsList"
        Me.grid_PartsList.Rows.Count = 2
        Me.grid_PartsList.Rows.DefaultSize = 20
        Me.grid_PartsList.Size = New System.Drawing.Size(568, 399)
        Me.grid_PartsList.StyleInfo = resources.GetString("grid_PartsList.StyleInfo")
        Me.grid_PartsList.TabIndex = 3
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Panel3.Controls.Add(Me.CheckBox1)
        Me.Panel3.Controls.Add(Me.Panel2)
        Me.Panel3.Controls.Add(Me.tb_Barcode)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1126, 307)
        Me.Panel3.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label39)
        Me.Panel2.Controls.Add(Me.tb_Option)
        Me.Panel2.Controls.Add(Me.Label22)
        Me.Panel2.Controls.Add(Me.tb_Lot_Status)
        Me.Panel2.Controls.Add(Me.GroupBox4)
        Me.Panel2.Controls.Add(Me.tb_ORG_Product)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.tb_Lot_Qty)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.tb_YJ_No)
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.tb_Product)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.tb_Lot_No)
        Me.Panel2.Location = New System.Drawing.Point(10, 65)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1104, 230)
        Me.Panel2.TabIndex = 26
        '
        'Label39
        '
        Me.Label39.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label39.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label39.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label39.ForeColor = System.Drawing.Color.White
        Me.Label39.Location = New System.Drawing.Point(607, 60)
        Me.Label39.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(120, 21)
        Me.Label39.TabIndex = 29
        Me.Label39.Text = "Option"
        Me.Label39.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Option
        '
        Me.tb_Option.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Option.Enabled = False
        Me.tb_Option.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Option.Location = New System.Drawing.Point(727, 60)
        Me.tb_Option.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Option.Name = "tb_Option"
        Me.tb_Option.Size = New System.Drawing.Size(180, 21)
        Me.tb_Option.TabIndex = 30
        '
        'Label22
        '
        Me.Label22.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label22.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label22.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.White
        Me.Label22.Location = New System.Drawing.Point(307, 60)
        Me.Label22.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(120, 21)
        Me.Label22.TabIndex = 27
        Me.Label22.Text = "Lot Status"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Lot_Status
        '
        Me.tb_Lot_Status.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Lot_Status.Enabled = False
        Me.tb_Lot_Status.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Lot_Status.Location = New System.Drawing.Point(427, 60)
        Me.tb_Lot_Status.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Lot_Status.Name = "tb_Lot_Status"
        Me.tb_Lot_Status.Size = New System.Drawing.Size(180, 21)
        Me.tb_Lot_Status.TabIndex = 28
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label14)
        Me.GroupBox4.Controls.Add(Me.Label13)
        Me.GroupBox4.Controls.Add(Me.tb_After_RCD)
        Me.GroupBox4.Controls.Add(Me.tb_Before_RCD)
        Me.GroupBox4.Controls.Add(Me.Label38)
        Me.GroupBox4.Controls.Add(Me.tb_After_PMIC)
        Me.GroupBox4.Controls.Add(Me.tb_Before_PMIC)
        Me.GroupBox4.Controls.Add(Me.Label37)
        Me.GroupBox4.Controls.Add(Me.tb_After_CC)
        Me.GroupBox4.Controls.Add(Me.tb_before_CC)
        Me.GroupBox4.Controls.Add(Me.Label36)
        Me.GroupBox4.Controls.Add(Me.Label35)
        Me.GroupBox4.Controls.Add(Me.Label25)
        Me.GroupBox4.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox4.ForeColor = System.Drawing.Color.White
        Me.GroupBox4.Location = New System.Drawing.Point(7, 95)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(559, 124)
        Me.GroupBox4.TabIndex = 26
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "수리정보"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(417, 91)
        Me.Label14.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(134, 21)
        Me.Label14.TabIndex = 28
        Me.Label14.Text = "수리진행"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.White
        Me.Label13.Location = New System.Drawing.Point(283, 91)
        Me.Label13.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(134, 21)
        Me.Label13.TabIndex = 27
        Me.Label13.Text = "수리진행"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_After_RCD
        '
        Me.tb_After_RCD.BackColor = System.Drawing.SystemColors.Window
        Me.tb_After_RCD.Enabled = False
        Me.tb_After_RCD.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_After_RCD.Location = New System.Drawing.Point(417, 68)
        Me.tb_After_RCD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_After_RCD.Name = "tb_After_RCD"
        Me.tb_After_RCD.Size = New System.Drawing.Size(134, 21)
        Me.tb_After_RCD.TabIndex = 5
        Me.tb_After_RCD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_Before_RCD
        '
        Me.tb_Before_RCD.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Before_RCD.Enabled = False
        Me.tb_Before_RCD.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Before_RCD.Location = New System.Drawing.Point(417, 45)
        Me.tb_Before_RCD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Before_RCD.Name = "tb_Before_RCD"
        Me.tb_Before_RCD.Size = New System.Drawing.Size(134, 21)
        Me.tb_Before_RCD.TabIndex = 4
        Me.tb_Before_RCD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label38
        '
        Me.Label38.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label38.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label38.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label38.ForeColor = System.Drawing.Color.White
        Me.Label38.Location = New System.Drawing.Point(417, 22)
        Me.Label38.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(134, 21)
        Me.Label38.TabIndex = 26
        Me.Label38.Text = "RCD"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_After_PMIC
        '
        Me.tb_After_PMIC.BackColor = System.Drawing.SystemColors.Window
        Me.tb_After_PMIC.Enabled = False
        Me.tb_After_PMIC.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_After_PMIC.Location = New System.Drawing.Point(283, 68)
        Me.tb_After_PMIC.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_After_PMIC.Name = "tb_After_PMIC"
        Me.tb_After_PMIC.Size = New System.Drawing.Size(134, 21)
        Me.tb_After_PMIC.TabIndex = 3
        Me.tb_After_PMIC.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_Before_PMIC
        '
        Me.tb_Before_PMIC.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Before_PMIC.Enabled = False
        Me.tb_Before_PMIC.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Before_PMIC.Location = New System.Drawing.Point(283, 45)
        Me.tb_Before_PMIC.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Before_PMIC.Name = "tb_Before_PMIC"
        Me.tb_Before_PMIC.Size = New System.Drawing.Size(134, 21)
        Me.tb_Before_PMIC.TabIndex = 2
        Me.tb_Before_PMIC.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label37
        '
        Me.Label37.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label37.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label37.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label37.ForeColor = System.Drawing.Color.White
        Me.Label37.Location = New System.Drawing.Point(283, 22)
        Me.Label37.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(134, 21)
        Me.Label37.TabIndex = 23
        Me.Label37.Text = "PMIC"
        Me.Label37.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_After_CC
        '
        Me.tb_After_CC.BackColor = System.Drawing.SystemColors.Window
        Me.tb_After_CC.Enabled = False
        Me.tb_After_CC.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_After_CC.Location = New System.Drawing.Point(149, 68)
        Me.tb_After_CC.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_After_CC.Name = "tb_After_CC"
        Me.tb_After_CC.Size = New System.Drawing.Size(134, 21)
        Me.tb_After_CC.TabIndex = 1
        Me.tb_After_CC.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_before_CC
        '
        Me.tb_before_CC.BackColor = System.Drawing.SystemColors.Window
        Me.tb_before_CC.Enabled = False
        Me.tb_before_CC.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_before_CC.Location = New System.Drawing.Point(149, 45)
        Me.tb_before_CC.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_before_CC.Name = "tb_before_CC"
        Me.tb_before_CC.Size = New System.Drawing.Size(134, 21)
        Me.tb_before_CC.TabIndex = 0
        Me.tb_before_CC.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label36
        '
        Me.Label36.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label36.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label36.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label36.ForeColor = System.Drawing.Color.White
        Me.Label36.Location = New System.Drawing.Point(149, 22)
        Me.Label36.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(134, 21)
        Me.Label36.TabIndex = 20
        Me.Label36.Text = "Customer Code"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label35
        '
        Me.Label35.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label35.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label35.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label35.ForeColor = System.Drawing.Color.White
        Me.Label35.Location = New System.Drawing.Point(15, 68)
        Me.Label35.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(134, 21)
        Me.Label35.TabIndex = 19
        Me.Label35.Text = "변경"
        Me.Label35.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label25.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label25.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.White
        Me.Label25.Location = New System.Drawing.Point(15, 45)
        Me.Label25.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(134, 21)
        Me.Label25.TabIndex = 18
        Me.Label25.Text = "기존"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_ORG_Product
        '
        Me.tb_ORG_Product.BackColor = System.Drawing.SystemColors.Window
        Me.tb_ORG_Product.Enabled = False
        Me.tb_ORG_Product.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_ORG_Product.Location = New System.Drawing.Point(127, 60)
        Me.tb_ORG_Product.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_ORG_Product.Name = "tb_ORG_Product"
        Me.tb_ORG_Product.Size = New System.Drawing.Size(180, 21)
        Me.tb_ORG_Product.TabIndex = 25
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(7, 60)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 21)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "기존 Product Code"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(836, 37)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 21)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Lot 수량"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Lot_Qty
        '
        Me.tb_Lot_Qty.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Lot_Qty.Enabled = False
        Me.tb_Lot_Qty.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Lot_Qty.Location = New System.Drawing.Point(956, 37)
        Me.tb_Lot_Qty.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Lot_Qty.Name = "tb_Lot_Qty"
        Me.tb_Lot_Qty.Size = New System.Drawing.Size(109, 21)
        Me.tb_Lot_Qty.TabIndex = 23
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(607, 37)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 21)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "YJ No."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_YJ_No
        '
        Me.tb_YJ_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_YJ_No.Enabled = False
        Me.tb_YJ_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_YJ_No.Location = New System.Drawing.Point(727, 37)
        Me.tb_YJ_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_YJ_No.Name = "tb_YJ_No"
        Me.tb_YJ_No.Size = New System.Drawing.Size(109, 21)
        Me.tb_YJ_No.TabIndex = 21
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold)
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(3, 6)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(190, 24)
        Me.Label10.TabIndex = 19
        Me.Label10.Text = "[ Lot Information ]"
        '
        'tb_Product
        '
        Me.tb_Product.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Product.Enabled = False
        Me.tb_Product.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Product.Location = New System.Drawing.Point(127, 37)
        Me.tb_Product.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Product.Name = "tb_Product"
        Me.tb_Product.Size = New System.Drawing.Size(180, 21)
        Me.tb_Product.TabIndex = 12
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(307, 37)
        Me.Label7.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 21)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Lot No."
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(7, 37)
        Me.Label6.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(120, 21)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "변경 Product Code"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Lot_No
        '
        Me.tb_Lot_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Lot_No.Enabled = False
        Me.tb_Lot_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Lot_No.Location = New System.Drawing.Point(427, 37)
        Me.tb_Lot_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Lot_No.Name = "tb_Lot_No"
        Me.tb_Lot_No.Size = New System.Drawing.Size(180, 21)
        Me.tb_Lot_No.TabIndex = 14
        '
        'tb_Barcode
        '
        Me.tb_Barcode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tb_Barcode.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Barcode.Enabled = False
        Me.tb_Barcode.Font = New System.Drawing.Font("굴림", 16.0!)
        Me.tb_Barcode.Location = New System.Drawing.Point(231, 10)
        Me.tb_Barcode.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Barcode.Name = "tb_Barcode"
        Me.tb_Barcode.Size = New System.Drawing.Size(886, 32)
        Me.tb_Barcode.TabIndex = 25
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Font = New System.Drawing.Font("굴림", 16.0!)
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(10, 10)
        Me.Label4.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(221, 32)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Module Barcode Scan"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(231, 46)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(90, 16)
        Me.CheckBox1.TabIndex = 29
        Me.CheckBox1.Text = "YJ No. 사용"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'frm_Parts_Kitting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1484, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_Parts_Kitting"
        Me.Text = "Kitting 등록"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.grid_KittingList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        CType(Me.grid_PartsList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents BTN_Search As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents btn_New_Kitting As ToolStripButton
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents btn_Save As ToolStripButton
    Friend WithEvents Form_CLose As ToolStripButton
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents btn_TestLabel As Button
    Friend WithEvents Label32 As Label
    Friend WithEvents Panel5 As Panel
    Friend WithEvents cb_PrinterList As ComboBox
    Friend WithEvents Label23 As Label
    Friend WithEvents CB_Print_Use As CheckBox
    Friend WithEvents TB_StopBits As TextBox
    Friend WithEvents Label34 As Label
    Friend WithEvents TB_Parity As TextBox
    Friend WithEvents Label33 As Label
    Friend WithEvents cb_Cable As ComboBox
    Friend WithEvents TB_DataBits As TextBox
    Friend WithEvents Label26 As Label
    Friend WithEvents Label27 As Label
    Friend WithEvents Label28 As Label
    Friend WithEvents TB_TOP_Loc As TextBox
    Friend WithEvents Label29 As Label
    Friend WithEvents TB_BaudRate As TextBox
    Friend WithEvents TB_LEFT_Loc As TextBox
    Friend WithEvents Label30 As Label
    Friend WithEvents Label31 As Label
    Friend WithEvents TB_Port As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents grid_KittingList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents tb_S_YJ_No As TextBox
    Friend WithEvents tb_S_Lot_No As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents DateTimePicker2 As DateTimePicker
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents Label11 As Label
    Friend WithEvents Panel3 As Panel
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents tb_After_RCD As TextBox
    Friend WithEvents tb_Before_RCD As TextBox
    Friend WithEvents Label38 As Label
    Friend WithEvents tb_After_PMIC As TextBox
    Friend WithEvents tb_Before_PMIC As TextBox
    Friend WithEvents Label37 As Label
    Friend WithEvents tb_After_CC As TextBox
    Friend WithEvents tb_before_CC As TextBox
    Friend WithEvents Label36 As Label
    Friend WithEvents Label35 As Label
    Friend WithEvents Label25 As Label
    Friend WithEvents tb_ORG_Product As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents tb_Lot_Qty As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents tb_YJ_No As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents tb_Product As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents tb_Lot_No As TextBox
    Friend WithEvents tb_Barcode As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents grid_PartsList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Label12 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents btn_Save2 As Button
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Label21 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents tb_Parts_Qty As TextBox
    Friend WithEvents tb_Parts_Barcode As TextBox
    Friend WithEvents Label20 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents tb_Parts_YJ_No As TextBox
    Friend WithEvents tb_Parts_PartNo As TextBox
    Friend WithEvents Label19 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents tb_Parts_Lot_No As TextBox
    Friend WithEvents btn_Check_Print As Button
    Friend WithEvents Label22 As Label
    Friend WithEvents tb_Lot_Status As TextBox
    Friend WithEvents Label24 As Label
    Friend WithEvents Label39 As Label
    Friend WithEvents tb_Option As TextBox
    Friend WithEvents CheckBox1 As CheckBox
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_LabelPrinter
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.cb_PrinterList = New System.Windows.Forms.ComboBox()
        Me.Label23 = New System.Windows.Forms.Label()
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.RadioButton10 = New System.Windows.Forms.RadioButton()
        Me.PictureBox6 = New System.Windows.Forms.PictureBox()
        Me.RadioButton9 = New System.Windows.Forms.RadioButton()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.RadioButton8 = New System.Windows.Forms.RadioButton()
        Me.btn_LabelPrint = New System.Windows.Forms.Button()
        Me.RadioButton7 = New System.Windows.Forms.RadioButton()
        Me.nud_PrintQty = New System.Windows.Forms.NumericUpDown()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RadioButton6 = New System.Windows.Forms.RadioButton()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.RadioButton5 = New System.Windows.Forms.RadioButton()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.pb_HeatCount2_Label = New System.Windows.Forms.PictureBox()
        Me.pb_HeatCount1_Label = New System.Windows.Forms.PictureBox()
        Me.pb_RCD_Label = New System.Windows.Forms.PictureBox()
        Me.pb_PMIC_Label = New System.Windows.Forms.PictureBox()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nud_PrintQty, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pb_HeatCount2_Label, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pb_HeatCount1_Label, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pb_RCD_Label, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pb_PMIC_Label, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label32)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel5)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.White
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 748)
        Me.SplitContainer1.SplitterDistance = 421
        Me.SplitContainer1.TabIndex = 0
        '
        'Label32
        '
        Me.Label32.BackColor = System.Drawing.Color.White
        Me.Label32.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label32.Font = New System.Drawing.Font("굴림", 16.0!, System.Drawing.FontStyle.Bold)
        Me.Label32.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label32.Location = New System.Drawing.Point(31, 10)
        Me.Label32.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(345, 47)
        Me.Label32.TabIndex = 3
        Me.Label32.Text = "프린터 설정"
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel5
        '
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel5.Controls.Add(Me.cb_PrinterList)
        Me.Panel5.Controls.Add(Me.Label23)
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
        Me.Panel5.Location = New System.Drawing.Point(31, 61)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(345, 268)
        Me.Panel5.TabIndex = 4
        '
        'cb_PrinterList
        '
        Me.cb_PrinterList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_PrinterList.Font = New System.Drawing.Font("굴림", 13.75!, System.Drawing.FontStyle.Bold)
        Me.cb_PrinterList.FormattingEnabled = True
        Me.cb_PrinterList.Location = New System.Drawing.Point(180, 35)
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
        Me.Label23.Location = New System.Drawing.Point(9, 35)
        Me.Label23.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(171, 26)
        Me.Label23.TabIndex = 3
        Me.Label23.Text = "Printer 목록"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_StopBits
        '
        Me.TB_StopBits.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_StopBits.Location = New System.Drawing.Point(180, 231)
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
        Me.Label34.Location = New System.Drawing.Point(9, 231)
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
        Me.TB_Parity.Location = New System.Drawing.Point(180, 203)
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
        Me.Label33.Location = New System.Drawing.Point(9, 203)
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
        Me.cb_Cable.Location = New System.Drawing.Point(180, 7)
        Me.cb_Cable.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.cb_Cable.Name = "cb_Cable"
        Me.cb_Cable.Size = New System.Drawing.Size(154, 26)
        Me.cb_Cable.TabIndex = 2
        '
        'TB_DataBits
        '
        Me.TB_DataBits.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_DataBits.Location = New System.Drawing.Point(180, 175)
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
        Me.Label26.Location = New System.Drawing.Point(9, 7)
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
        Me.Label27.Location = New System.Drawing.Point(9, 175)
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
        Me.Label28.Location = New System.Drawing.Point(9, 119)
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
        Me.TB_TOP_Loc.Location = New System.Drawing.Point(180, 119)
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
        Me.Label29.Location = New System.Drawing.Point(9, 147)
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
        Me.TB_BaudRate.Location = New System.Drawing.Point(180, 147)
        Me.TB_BaudRate.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_BaudRate.Name = "TB_BaudRate"
        Me.TB_BaudRate.Size = New System.Drawing.Size(154, 26)
        Me.TB_BaudRate.TabIndex = 12
        Me.TB_BaudRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TB_LEFT_Loc
        '
        Me.TB_LEFT_Loc.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_LEFT_Loc.Location = New System.Drawing.Point(180, 91)
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
        Me.Label30.Location = New System.Drawing.Point(9, 63)
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
        Me.Label31.Location = New System.Drawing.Point(9, 91)
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
        Me.TB_Port.Location = New System.Drawing.Point(180, 63)
        Me.TB_Port.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Port.Name = "TB_Port"
        Me.TB_Port.Size = New System.Drawing.Size(154, 26)
        Me.TB_Port.TabIndex = 6
        Me.TB_Port.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.RadioButton10)
        Me.GroupBox1.Controls.Add(Me.PictureBox6)
        Me.GroupBox1.Controls.Add(Me.RadioButton9)
        Me.GroupBox1.Controls.Add(Me.PictureBox5)
        Me.GroupBox1.Controls.Add(Me.RadioButton8)
        Me.GroupBox1.Controls.Add(Me.btn_LabelPrint)
        Me.GroupBox1.Controls.Add(Me.RadioButton7)
        Me.GroupBox1.Controls.Add(Me.nud_PrintQty)
        Me.GroupBox1.Controls.Add(Me.PictureBox3)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.RadioButton6)
        Me.GroupBox1.Controls.Add(Me.PictureBox2)
        Me.GroupBox1.Controls.Add(Me.RadioButton5)
        Me.GroupBox1.Controls.Add(Me.PictureBox1)
        Me.GroupBox1.Controls.Add(Me.RadioButton4)
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Controls.Add(Me.pb_HeatCount2_Label)
        Me.GroupBox1.Controls.Add(Me.pb_HeatCount1_Label)
        Me.GroupBox1.Controls.Add(Me.pb_RCD_Label)
        Me.GroupBox1.Controls.Add(Me.pb_PMIC_Label)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(839, 748)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "발행라벨 선택"
        '
        'Label2
        '
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.Location = New System.Drawing.Point(380, 344)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(305, 93)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "   N/V"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'RadioButton10
        '
        Me.RadioButton10.AutoSize = True
        Me.RadioButton10.Location = New System.Drawing.Point(360, 483)
        Me.RadioButton10.Name = "RadioButton10"
        Me.RadioButton10.Size = New System.Drawing.Size(14, 13)
        Me.RadioButton10.TabIndex = 19
        Me.RadioButton10.TabStop = True
        Me.RadioButton10.UseVisualStyleBackColor = True
        Me.RadioButton10.Visible = False
        '
        'PictureBox6
        '
        Me.PictureBox6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox6.Location = New System.Drawing.Point(380, 443)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(305, 93)
        Me.PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox6.TabIndex = 18
        Me.PictureBox6.TabStop = False
        Me.PictureBox6.Visible = False
        '
        'RadioButton9
        '
        Me.RadioButton9.AutoSize = True
        Me.RadioButton9.Location = New System.Drawing.Point(29, 483)
        Me.RadioButton9.Name = "RadioButton9"
        Me.RadioButton9.Size = New System.Drawing.Size(14, 13)
        Me.RadioButton9.TabIndex = 17
        Me.RadioButton9.TabStop = True
        Me.RadioButton9.UseVisualStyleBackColor = True
        '
        'PictureBox5
        '
        Me.PictureBox5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox5.Image = Global.Repair_System.My.Resources.Resources.RCD_Count
        Me.PictureBox5.Location = New System.Drawing.Point(49, 443)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(305, 93)
        Me.PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox5.TabIndex = 16
        Me.PictureBox5.TabStop = False
        '
        'RadioButton8
        '
        Me.RadioButton8.AutoSize = True
        Me.RadioButton8.Location = New System.Drawing.Point(360, 384)
        Me.RadioButton8.Name = "RadioButton8"
        Me.RadioButton8.Size = New System.Drawing.Size(14, 13)
        Me.RadioButton8.TabIndex = 15
        Me.RadioButton8.TabStop = True
        Me.RadioButton8.UseVisualStyleBackColor = True
        '
        'btn_LabelPrint
        '
        Me.btn_LabelPrint.Font = New System.Drawing.Font("굴림", 15.75!, System.Drawing.FontStyle.Bold)
        Me.btn_LabelPrint.Location = New System.Drawing.Point(212, 624)
        Me.btn_LabelPrint.Name = "btn_LabelPrint"
        Me.btn_LabelPrint.Size = New System.Drawing.Size(141, 82)
        Me.btn_LabelPrint.TabIndex = 3
        Me.btn_LabelPrint.Text = "라벨 발행"
        Me.btn_LabelPrint.UseVisualStyleBackColor = True
        '
        'RadioButton7
        '
        Me.RadioButton7.AutoSize = True
        Me.RadioButton7.Location = New System.Drawing.Point(360, 285)
        Me.RadioButton7.Name = "RadioButton7"
        Me.RadioButton7.Size = New System.Drawing.Size(14, 13)
        Me.RadioButton7.TabIndex = 13
        Me.RadioButton7.TabStop = True
        Me.RadioButton7.UseVisualStyleBackColor = True
        '
        'nud_PrintQty
        '
        Me.nud_PrintQty.Font = New System.Drawing.Font("굴림", 15.75!, System.Drawing.FontStyle.Bold)
        Me.nud_PrintQty.Location = New System.Drawing.Point(127, 573)
        Me.nud_PrintQty.Maximum = New Decimal(New Integer() {5000, 0, 0, 0})
        Me.nud_PrintQty.Name = "nud_PrintQty"
        Me.nud_PrintQty.Size = New System.Drawing.Size(227, 32)
        Me.nud_PrintQty.TabIndex = 2
        Me.nud_PrintQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.nud_PrintQty.Value = New Decimal(New Integer() {4993, 0, 0, 0})
        '
        'PictureBox3
        '
        Me.PictureBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox3.Image = Global.Repair_System.My.Resources.Resources.PMIC_Reg_FAIL
        Me.PictureBox3.Location = New System.Drawing.Point(380, 245)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(305, 93)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 12
        Me.PictureBox3.TabStop = False
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.Location = New System.Drawing.Point(10, 573)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 32)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "발행 수"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton6
        '
        Me.RadioButton6.AutoSize = True
        Me.RadioButton6.Location = New System.Drawing.Point(360, 186)
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.Size = New System.Drawing.Size(14, 13)
        Me.RadioButton6.TabIndex = 11
        Me.RadioButton6.TabStop = True
        Me.RadioButton6.UseVisualStyleBackColor = True
        '
        'PictureBox2
        '
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Image = Global.Repair_System.My.Resources.Resources.Map_Out
        Me.PictureBox2.Location = New System.Drawing.Point(380, 146)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(305, 93)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 10
        Me.PictureBox2.TabStop = False
        '
        'RadioButton5
        '
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Location = New System.Drawing.Point(360, 87)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(14, 13)
        Me.RadioButton5.TabIndex = 9
        Me.RadioButton5.TabStop = True
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = Global.Repair_System.My.Resources.Resources.SRB_CRB_Test_Fail
        Me.PictureBox1.Location = New System.Drawing.Point(380, 47)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(305, 93)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 8
        Me.PictureBox1.TabStop = False
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Location = New System.Drawing.Point(29, 384)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(14, 13)
        Me.RadioButton4.TabIndex = 7
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(29, 285)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(14, 13)
        Me.RadioButton3.TabIndex = 6
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(29, 186)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(14, 13)
        Me.RadioButton2.TabIndex = 5
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(29, 87)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(14, 13)
        Me.RadioButton1.TabIndex = 4
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'pb_HeatCount2_Label
        '
        Me.pb_HeatCount2_Label.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pb_HeatCount2_Label.Image = Global.Repair_System.My.Resources.Resources.HearCount2
        Me.pb_HeatCount2_Label.Location = New System.Drawing.Point(49, 344)
        Me.pb_HeatCount2_Label.Name = "pb_HeatCount2_Label"
        Me.pb_HeatCount2_Label.Size = New System.Drawing.Size(305, 93)
        Me.pb_HeatCount2_Label.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pb_HeatCount2_Label.TabIndex = 3
        Me.pb_HeatCount2_Label.TabStop = False
        '
        'pb_HeatCount1_Label
        '
        Me.pb_HeatCount1_Label.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pb_HeatCount1_Label.Image = Global.Repair_System.My.Resources.Resources.HearCount1
        Me.pb_HeatCount1_Label.Location = New System.Drawing.Point(49, 245)
        Me.pb_HeatCount1_Label.Name = "pb_HeatCount1_Label"
        Me.pb_HeatCount1_Label.Size = New System.Drawing.Size(305, 93)
        Me.pb_HeatCount1_Label.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pb_HeatCount1_Label.TabIndex = 2
        Me.pb_HeatCount1_Label.TabStop = False
        '
        'pb_RCD_Label
        '
        Me.pb_RCD_Label.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pb_RCD_Label.Image = Global.Repair_System.My.Resources.Resources.RCD
        Me.pb_RCD_Label.Location = New System.Drawing.Point(49, 146)
        Me.pb_RCD_Label.Name = "pb_RCD_Label"
        Me.pb_RCD_Label.Size = New System.Drawing.Size(305, 93)
        Me.pb_RCD_Label.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pb_RCD_Label.TabIndex = 1
        Me.pb_RCD_Label.TabStop = False
        '
        'pb_PMIC_Label
        '
        Me.pb_PMIC_Label.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pb_PMIC_Label.Image = Global.Repair_System.My.Resources.Resources.PMIC
        Me.pb_PMIC_Label.Location = New System.Drawing.Point(49, 47)
        Me.pb_PMIC_Label.Name = "pb_PMIC_Label"
        Me.pb_PMIC_Label.Size = New System.Drawing.Size(305, 93)
        Me.pb_PMIC_Label.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pb_PMIC_Label.TabIndex = 0
        Me.pb_PMIC_Label.TabStop = False
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Form_CLose})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1264, 25)
        Me.ToolStrip1.TabIndex = 17
        Me.ToolStrip1.Text = "ToolStrip1"
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
        'frm_LabelPrinter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_LabelPrinter"
        Me.Text = "각종 라벨 발행"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nud_PrintQty, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pb_HeatCount2_Label, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pb_HeatCount1_Label, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pb_RCD_Label, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pb_PMIC_Label, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Label32 As Label
    Friend WithEvents Panel5 As Panel
    Friend WithEvents cb_PrinterList As ComboBox
    Friend WithEvents Label23 As Label
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
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents pb_PMIC_Label As PictureBox
    Friend WithEvents pb_RCD_Label As PictureBox
    Friend WithEvents pb_HeatCount2_Label As PictureBox
    Friend WithEvents pb_HeatCount1_Label As PictureBox
    Friend WithEvents RadioButton4 As RadioButton
    Friend WithEvents RadioButton3 As RadioButton
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents btn_LabelPrint As Button
    Friend WithEvents nud_PrintQty As NumericUpDown
    Friend WithEvents Label1 As Label
    Friend WithEvents RadioButton7 As RadioButton
    Friend WithEvents PictureBox3 As PictureBox
    Friend WithEvents RadioButton6 As RadioButton
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents RadioButton5 As RadioButton
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents RadioButton8 As RadioButton
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents Form_CLose As ToolStripButton
    Friend WithEvents RadioButton9 As RadioButton
    Friend WithEvents PictureBox5 As PictureBox
    Friend WithEvents RadioButton10 As RadioButton
    Friend WithEvents PictureBox6 As PictureBox
    Friend WithEvents Label2 As Label
End Class

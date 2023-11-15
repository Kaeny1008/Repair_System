<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BluModelManager
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
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BluModelManager))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.RB_Model_Use = New System.Windows.Forms.RadioButton()
        Me.RB_Model_Not_Use = New System.Windows.Forms.RadioButton()
        Me.RB_Model_ALL = New System.Windows.Forms.RadioButton()
        Me.CB_Model_Division = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TB_Model_Name = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.GRID_ModelList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.OFP_DELETE = New System.Windows.Forms.Button()
        Me.OFP_DOWNLOAD = New System.Windows.Forms.Button()
        Me.OQC_FLOOR_PLAN_FILE = New System.Windows.Forms.TextBox()
        Me.OFP_UPLOAD = New System.Windows.Forms.Button()
        Me.OQC_FLOOR_PLAN_DATE = New System.Windows.Forms.DateTimePicker()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.MODEL_PART_NO = New System.Windows.Forms.TextBox()
        Me.PACKING_QTY = New System.Windows.Forms.NumericUpDown()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.CHROMATICITY_RANK = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.INTENSITY_RANK = New System.Windows.Forms.TextBox()
        Me.TH_QTY = New System.Windows.Forms.NumericUpDown()
        Me.PKG_QTY = New System.Windows.Forms.NumericUpDown()
        Me.PCB_PARTNO = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.RB_NOT_USE = New System.Windows.Forms.RadioButton()
        Me.RB_USE = New System.Windows.Forms.RadioButton()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.MODEL_NOTE = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.TH_PARTNO = New System.Windows.Forms.TextBox()
        Me.CB_TH_Maker = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.CB_RankOrder = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.PKG_PARTNO = New System.Windows.Forms.TextBox()
        Me.CB_PKG_Maker = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.CB_ModelDivision = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.MARKING_DELETE = New System.Windows.Forms.Button()
        Me.FLOOR_PLAN_DELETE = New System.Windows.Forms.Button()
        Me.APP_SHEET_DELETE = New System.Windows.Forms.Button()
        Me.MARKING_DOWNLOAD = New System.Windows.Forms.Button()
        Me.MARKING_FILE = New System.Windows.Forms.TextBox()
        Me.MARKING_UPLOAD = New System.Windows.Forms.Button()
        Me.MARKING_DATE = New System.Windows.Forms.DateTimePicker()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.FLOOR_PLAN_DOWNLOAD = New System.Windows.Forms.Button()
        Me.FLOOR_PLAN_FILE = New System.Windows.Forms.TextBox()
        Me.FLOOR_PLAN_UPLOAD = New System.Windows.Forms.Button()
        Me.FLOOR_PLAN_DATE = New System.Windows.Forms.DateTimePicker()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.APP_SHEET_DOWNLOAD = New System.Windows.Forms.Button()
        Me.APP_SHEET_FILE = New System.Windows.Forms.TextBox()
        Me.APP_SHEET_UPLOAD = New System.Windows.Forms.Button()
        Me.APP_SHEET_DATE = New System.Windows.Forms.DateTimePicker()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.MODEL_NAME = New System.Windows.Forms.TextBox()
        Me.MODEL_CODE = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.BTN_Search = New System.Windows.Forms.ToolStripButton()
        Me.BTN_NewModel = New System.Windows.Forms.ToolStripButton()
        Me.BTN_NewCancel = New System.Windows.Forms.ToolStripButton()
        Me.BTN_ModelSave = New System.Windows.Forms.ToolStripButton()
        Me.GRID_MENU = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.BTN_ADD = New System.Windows.Forms.ToolStripMenuItem()
        Me.BTN_DELETE = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_Save = New System.Windows.Forms.ToolStripMenuItem()
        Me.GRID_MENU2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.BTN_ROW_ADD = New System.Windows.Forms.ToolStripMenuItem()
        Me.BTN_ROW_DELETE = New System.Windows.Forms.ToolStripMenuItem()
        Me.BTN_COL_ADD = New System.Windows.Forms.ToolStripMenuItem()
        Me.BTN_COL_DELETE = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.GRID_ModelList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PACKING_QTY, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TH_QTY, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PKG_QTY, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.GRID_MENU.SuspendLayout()
        Me.GRID_MENU2.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RB_Model_Use)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RB_Model_Not_Use)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RB_Model_ALL)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CB_Model_Division)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TB_Model_Name)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1394, 734)
        Me.SplitContainer1.SplitterDistance = 43
        Me.SplitContainer1.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(9, 11)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 21)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "모델상태"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RB_Model_Use
        '
        Me.RB_Model_Use.AutoSize = True
        Me.RB_Model_Use.ForeColor = System.Drawing.Color.White
        Me.RB_Model_Use.Location = New System.Drawing.Point(226, 13)
        Me.RB_Model_Use.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.RB_Model_Use.Name = "RB_Model_Use"
        Me.RB_Model_Use.Size = New System.Drawing.Size(47, 16)
        Me.RB_Model_Use.TabIndex = 7
        Me.RB_Model_Use.Text = "사용"
        Me.RB_Model_Use.UseVisualStyleBackColor = True
        '
        'RB_Model_Not_Use
        '
        Me.RB_Model_Not_Use.AutoSize = True
        Me.RB_Model_Not_Use.ForeColor = System.Drawing.Color.White
        Me.RB_Model_Not_Use.Location = New System.Drawing.Point(159, 13)
        Me.RB_Model_Not_Use.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.RB_Model_Not_Use.Name = "RB_Model_Not_Use"
        Me.RB_Model_Not_Use.Size = New System.Drawing.Size(59, 16)
        Me.RB_Model_Not_Use.TabIndex = 6
        Me.RB_Model_Not_Use.Text = "미사용"
        Me.RB_Model_Not_Use.UseVisualStyleBackColor = True
        '
        'RB_Model_ALL
        '
        Me.RB_Model_ALL.AutoSize = True
        Me.RB_Model_ALL.Checked = True
        Me.RB_Model_ALL.ForeColor = System.Drawing.Color.White
        Me.RB_Model_ALL.Location = New System.Drawing.Point(104, 13)
        Me.RB_Model_ALL.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.RB_Model_ALL.Name = "RB_Model_ALL"
        Me.RB_Model_ALL.Size = New System.Drawing.Size(47, 16)
        Me.RB_Model_ALL.TabIndex = 5
        Me.RB_Model_ALL.TabStop = True
        Me.RB_Model_ALL.Text = "모두"
        Me.RB_Model_ALL.UseVisualStyleBackColor = True
        '
        'CB_Model_Division
        '
        Me.CB_Model_Division.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_Model_Division.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.CB_Model_Division.FormattingEnabled = True
        Me.CB_Model_Division.Items.AddRange(New Object() {"개발", "양산", "모두"})
        Me.CB_Model_Division.Location = New System.Drawing.Point(359, 11)
        Me.CB_Model_Division.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.CB_Model_Division.Name = "CB_Model_Division"
        Me.CB_Model_Division.Size = New System.Drawing.Size(189, 21)
        Me.CB_Model_Division.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(273, 11)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 21)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "개발 / 양산"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TB_Model_Name
        '
        Me.TB_Model_Name.Location = New System.Drawing.Point(749, 11)
        Me.TB_Model_Name.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Model_Name.Name = "TB_Model_Name"
        Me.TB_Model_Name.Size = New System.Drawing.Size(254, 21)
        Me.TB_Model_Name.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(549, 11)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(200, 21)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Model CODE / Model Name"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.GRID_ModelList)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.BackColor = System.Drawing.Color.AliceBlue
        Me.SplitContainer2.Panel2.Controls.Add(Me.OFP_DELETE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.OFP_DOWNLOAD)
        Me.SplitContainer2.Panel2.Controls.Add(Me.OQC_FLOOR_PLAN_FILE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.OFP_UPLOAD)
        Me.SplitContainer2.Panel2.Controls.Add(Me.OQC_FLOOR_PLAN_DATE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label20)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label16)
        Me.SplitContainer2.Panel2.Controls.Add(Me.MODEL_PART_NO)
        Me.SplitContainer2.Panel2.Controls.Add(Me.PACKING_QTY)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label14)
        Me.SplitContainer2.Panel2.Controls.Add(Me.ListBox1)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer2.Panel2.Controls.Add(Me.CHROMATICITY_RANK)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel2.Controls.Add(Me.INTENSITY_RANK)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TH_QTY)
        Me.SplitContainer2.Panel2.Controls.Add(Me.PKG_QTY)
        Me.SplitContainer2.Panel2.Controls.Add(Me.PCB_PARTNO)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label19)
        Me.SplitContainer2.Panel2.Controls.Add(Me.RB_NOT_USE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.RB_USE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label18)
        Me.SplitContainer2.Panel2.Controls.Add(Me.MODEL_NOTE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label17)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TH_PARTNO)
        Me.SplitContainer2.Panel2.Controls.Add(Me.CB_TH_Maker)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label13)
        Me.SplitContainer2.Panel2.Controls.Add(Me.CB_RankOrder)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer2.Panel2.Controls.Add(Me.PKG_PARTNO)
        Me.SplitContainer2.Panel2.Controls.Add(Me.CB_PKG_Maker)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer2.Panel2.Controls.Add(Me.CB_ModelDivision)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer2.Panel2.Controls.Add(Me.MARKING_DELETE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.FLOOR_PLAN_DELETE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.APP_SHEET_DELETE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.MARKING_DOWNLOAD)
        Me.SplitContainer2.Panel2.Controls.Add(Me.MARKING_FILE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.MARKING_UPLOAD)
        Me.SplitContainer2.Panel2.Controls.Add(Me.MARKING_DATE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer2.Panel2.Controls.Add(Me.FLOOR_PLAN_DOWNLOAD)
        Me.SplitContainer2.Panel2.Controls.Add(Me.FLOOR_PLAN_FILE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.FLOOR_PLAN_UPLOAD)
        Me.SplitContainer2.Panel2.Controls.Add(Me.FLOOR_PLAN_DATE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.APP_SHEET_DOWNLOAD)
        Me.SplitContainer2.Panel2.Controls.Add(Me.APP_SHEET_FILE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.APP_SHEET_UPLOAD)
        Me.SplitContainer2.Panel2.Controls.Add(Me.APP_SHEET_DATE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer2.Panel2.Controls.Add(Me.MODEL_NAME)
        Me.SplitContainer2.Panel2.Controls.Add(Me.MODEL_CODE)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer2.Size = New System.Drawing.Size(1394, 687)
        Me.SplitContainer2.SplitterDistance = 98
        Me.SplitContainer2.TabIndex = 0
        '
        'GRID_ModelList
        '
        Me.GRID_ModelList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.GRID_ModelList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GRID_ModelList.Location = New System.Drawing.Point(0, 0)
        Me.GRID_ModelList.Name = "GRID_ModelList"
        Me.GRID_ModelList.Rows.Count = 2
        Me.GRID_ModelList.Rows.DefaultSize = 20
        Me.GRID_ModelList.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row
        Me.GRID_ModelList.Size = New System.Drawing.Size(1394, 98)
        Me.GRID_ModelList.StyleInfo = resources.GetString("GRID_ModelList.StyleInfo")
        Me.GRID_ModelList.TabIndex = 0
        Me.GRID_ModelList.VisualStyle = C1.Win.C1FlexGrid.VisualStyle.Office2010Blue
        '
        'OFP_DELETE
        '
        Me.OFP_DELETE.Location = New System.Drawing.Point(814, 157)
        Me.OFP_DELETE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.OFP_DELETE.Name = "OFP_DELETE"
        Me.OFP_DELETE.Size = New System.Drawing.Size(75, 23)
        Me.OFP_DELETE.TabIndex = 55
        Me.OFP_DELETE.Text = "Delete"
        Me.OFP_DELETE.UseVisualStyleBackColor = True
        '
        'OFP_DOWNLOAD
        '
        Me.OFP_DOWNLOAD.Location = New System.Drawing.Point(739, 157)
        Me.OFP_DOWNLOAD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.OFP_DOWNLOAD.Name = "OFP_DOWNLOAD"
        Me.OFP_DOWNLOAD.Size = New System.Drawing.Size(75, 23)
        Me.OFP_DOWNLOAD.TabIndex = 54
        Me.OFP_DOWNLOAD.Text = "Download"
        Me.OFP_DOWNLOAD.UseVisualStyleBackColor = True
        '
        'OQC_FLOOR_PLAN_FILE
        '
        Me.OQC_FLOOR_PLAN_FILE.BackColor = System.Drawing.SystemColors.Window
        Me.OQC_FLOOR_PLAN_FILE.Enabled = False
        Me.OQC_FLOOR_PLAN_FILE.Location = New System.Drawing.Point(332, 158)
        Me.OQC_FLOOR_PLAN_FILE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.OQC_FLOOR_PLAN_FILE.Name = "OQC_FLOOR_PLAN_FILE"
        Me.OQC_FLOOR_PLAN_FILE.Size = New System.Drawing.Size(329, 21)
        Me.OQC_FLOOR_PLAN_FILE.TabIndex = 52
        Me.OQC_FLOOR_PLAN_FILE.Text = "등록안됨"
        '
        'OFP_UPLOAD
        '
        Me.OFP_UPLOAD.Location = New System.Drawing.Point(664, 157)
        Me.OFP_UPLOAD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.OFP_UPLOAD.Name = "OFP_UPLOAD"
        Me.OFP_UPLOAD.Size = New System.Drawing.Size(75, 23)
        Me.OFP_UPLOAD.TabIndex = 53
        Me.OFP_UPLOAD.Text = "Upload"
        Me.OFP_UPLOAD.UseVisualStyleBackColor = True
        '
        'OQC_FLOOR_PLAN_DATE
        '
        Me.OQC_FLOOR_PLAN_DATE.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.OQC_FLOOR_PLAN_DATE.Location = New System.Drawing.Point(229, 158)
        Me.OQC_FLOOR_PLAN_DATE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.OQC_FLOOR_PLAN_DATE.Name = "OQC_FLOOR_PLAN_DATE"
        Me.OQC_FLOOR_PLAN_DATE.Size = New System.Drawing.Size(103, 21)
        Me.OQC_FLOOR_PLAN_DATE.TabIndex = 51
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label20.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label20.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label20.ForeColor = System.Drawing.Color.White
        Me.Label20.Location = New System.Drawing.Point(9, 158)
        Me.Label20.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(220, 21)
        Me.Label20.TabIndex = 50
        Me.Label20.Text = "도해(출하성적서)"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(892, 144)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(457, 12)
        Me.Label16.TabIndex = 49
        Me.Label16.Text = "※ 승인원, 도면, 마킹기준, 도해 파일이 등록되지 않으면 날짜도 등록되지 않습니다."
        '
        'MODEL_PART_NO
        '
        Me.MODEL_PART_NO.BackColor = System.Drawing.SystemColors.Window
        Me.MODEL_PART_NO.Location = New System.Drawing.Point(661, 55)
        Me.MODEL_PART_NO.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.MODEL_PART_NO.Name = "MODEL_PART_NO"
        Me.MODEL_PART_NO.Size = New System.Drawing.Size(228, 21)
        Me.MODEL_PART_NO.TabIndex = 48
        '
        'PACKING_QTY
        '
        Me.PACKING_QTY.Location = New System.Drawing.Point(884, 236)
        Me.PACKING_QTY.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.PACKING_QTY.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.PACKING_QTY.Name = "PACKING_QTY"
        Me.PACKING_QTY.Size = New System.Drawing.Size(432, 21)
        Me.PACKING_QTY.TabIndex = 45
        Me.PACKING_QTY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label15.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(664, 236)
        Me.Label15.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(220, 21)
        Me.Label15.TabIndex = 44
        Me.Label15.Text = "최대 Packing수량"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label14.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(9, 283)
        Me.Label14.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(220, 208)
        Me.Label14.TabIndex = 36
        Me.Label14.Text = "Rank List"
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 12
        Me.ListBox1.Location = New System.Drawing.Point(229, 283)
        Me.ListBox1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(432, 208)
        Me.ListBox1.TabIndex = 37
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label12.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label12.ForeColor = System.Drawing.Color.White
        Me.Label12.Location = New System.Drawing.Point(9, 260)
        Me.Label12.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(220, 21)
        Me.Label12.TabIndex = 34
        Me.Label12.Text = "Chromaticity Rank(색좌표)"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CHROMATICITY_RANK
        '
        Me.CHROMATICITY_RANK.BackColor = System.Drawing.SystemColors.Window
        Me.CHROMATICITY_RANK.Location = New System.Drawing.Point(229, 260)
        Me.CHROMATICITY_RANK.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.CHROMATICITY_RANK.Name = "CHROMATICITY_RANK"
        Me.CHROMATICITY_RANK.Size = New System.Drawing.Size(432, 21)
        Me.CHROMATICITY_RANK.TabIndex = 35
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label11.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(9, 237)
        Me.Label11.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(220, 21)
        Me.Label11.TabIndex = 32
        Me.Label11.Text = "Intensity Rank(휘도)"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'INTENSITY_RANK
        '
        Me.INTENSITY_RANK.BackColor = System.Drawing.SystemColors.Window
        Me.INTENSITY_RANK.Location = New System.Drawing.Point(229, 237)
        Me.INTENSITY_RANK.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.INTENSITY_RANK.Name = "INTENSITY_RANK"
        Me.INTENSITY_RANK.Size = New System.Drawing.Size(432, 21)
        Me.INTENSITY_RANK.TabIndex = 33
        '
        'TH_QTY
        '
        Me.TH_QTY.Location = New System.Drawing.Point(1233, 191)
        Me.TH_QTY.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TH_QTY.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.TH_QTY.Name = "TH_QTY"
        Me.TH_QTY.Size = New System.Drawing.Size(83, 21)
        Me.TH_QTY.TabIndex = 41
        Me.TH_QTY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PKG_QTY
        '
        Me.PKG_QTY.Location = New System.Drawing.Point(572, 191)
        Me.PKG_QTY.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.PKG_QTY.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.PKG_QTY.Name = "PKG_QTY"
        Me.PKG_QTY.Size = New System.Drawing.Size(89, 21)
        Me.PKG_QTY.TabIndex = 29
        Me.PKG_QTY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PCB_PARTNO
        '
        Me.PCB_PARTNO.BackColor = System.Drawing.SystemColors.Window
        Me.PCB_PARTNO.Location = New System.Drawing.Point(884, 214)
        Me.PCB_PARTNO.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.PCB_PARTNO.Name = "PCB_PARTNO"
        Me.PCB_PARTNO.Size = New System.Drawing.Size(432, 21)
        Me.PCB_PARTNO.TabIndex = 43
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label19.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.White
        Me.Label19.Location = New System.Drawing.Point(664, 214)
        Me.Label19.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(220, 21)
        Me.Label19.TabIndex = 42
        Me.Label19.Text = "PCB Part No."
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'RB_NOT_USE
        '
        Me.RB_NOT_USE.AutoSize = True
        Me.RB_NOT_USE.Location = New System.Drawing.Point(287, 11)
        Me.RB_NOT_USE.Name = "RB_NOT_USE"
        Me.RB_NOT_USE.Size = New System.Drawing.Size(59, 16)
        Me.RB_NOT_USE.TabIndex = 2
        Me.RB_NOT_USE.Text = "미사용"
        Me.RB_NOT_USE.UseVisualStyleBackColor = True
        '
        'RB_USE
        '
        Me.RB_USE.AutoSize = True
        Me.RB_USE.Checked = True
        Me.RB_USE.Location = New System.Drawing.Point(232, 11)
        Me.RB_USE.Name = "RB_USE"
        Me.RB_USE.Size = New System.Drawing.Size(47, 16)
        Me.RB_USE.TabIndex = 1
        Me.RB_USE.TabStop = True
        Me.RB_USE.Text = "사용"
        Me.RB_USE.UseVisualStyleBackColor = True
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label18.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label18.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label18.ForeColor = System.Drawing.Color.White
        Me.Label18.Location = New System.Drawing.Point(9, 9)
        Me.Label18.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(220, 21)
        Me.Label18.TabIndex = 0
        Me.Label18.Text = "사용여부(생산중지)"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'MODEL_NOTE
        '
        Me.MODEL_NOTE.BackColor = System.Drawing.SystemColors.Window
        Me.MODEL_NOTE.Location = New System.Drawing.Point(884, 259)
        Me.MODEL_NOTE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.MODEL_NOTE.Multiline = True
        Me.MODEL_NOTE.Name = "MODEL_NOTE"
        Me.MODEL_NOTE.Size = New System.Drawing.Size(432, 95)
        Me.MODEL_NOTE.TabIndex = 0
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label17.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label17.ForeColor = System.Drawing.Color.White
        Me.Label17.Location = New System.Drawing.Point(664, 259)
        Me.Label17.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(220, 95)
        Me.Label17.TabIndex = 47
        Me.Label17.Text = "비고"
        '
        'TH_PARTNO
        '
        Me.TH_PARTNO.BackColor = System.Drawing.SystemColors.Window
        Me.TH_PARTNO.Location = New System.Drawing.Point(987, 191)
        Me.TH_PARTNO.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TH_PARTNO.Name = "TH_PARTNO"
        Me.TH_PARTNO.Size = New System.Drawing.Size(268, 21)
        Me.TH_PARTNO.TabIndex = 40
        '
        'CB_TH_Maker
        '
        Me.CB_TH_Maker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_TH_Maker.Font = New System.Drawing.Font("굴림", 9.5!)
        Me.CB_TH_Maker.FormattingEnabled = True
        Me.CB_TH_Maker.Location = New System.Drawing.Point(884, 191)
        Me.CB_TH_Maker.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.CB_TH_Maker.Name = "CB_TH_Maker"
        Me.CB_TH_Maker.Size = New System.Drawing.Size(103, 21)
        Me.CB_TH_Maker.TabIndex = 39
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label13.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label13.ForeColor = System.Drawing.Color.White
        Me.Label13.Location = New System.Drawing.Point(664, 191)
        Me.Label13.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(220, 21)
        Me.Label13.TabIndex = 38
        Me.Label13.Text = "TH (Maker, Part No, 사용수량)"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CB_RankOrder
        '
        Me.CB_RankOrder.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_RankOrder.Font = New System.Drawing.Font("굴림", 9.5!)
        Me.CB_RankOrder.FormattingEnabled = True
        Me.CB_RankOrder.Items.AddRange(New Object() {"휘도-색좌표", "색좌표-휘도"})
        Me.CB_RankOrder.Location = New System.Drawing.Point(229, 214)
        Me.CB_RankOrder.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.CB_RankOrder.Name = "CB_RankOrder"
        Me.CB_RankOrder.Size = New System.Drawing.Size(432, 21)
        Me.CB_RankOrder.TabIndex = 31
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label10.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(9, 214)
        Me.Label10.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(220, 21)
        Me.Label10.TabIndex = 30
        Me.Label10.Text = "RANK 순서"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PKG_PARTNO
        '
        Me.PKG_PARTNO.BackColor = System.Drawing.SystemColors.Window
        Me.PKG_PARTNO.Location = New System.Drawing.Point(332, 191)
        Me.PKG_PARTNO.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.PKG_PARTNO.Name = "PKG_PARTNO"
        Me.PKG_PARTNO.Size = New System.Drawing.Size(268, 21)
        Me.PKG_PARTNO.TabIndex = 28
        '
        'CB_PKG_Maker
        '
        Me.CB_PKG_Maker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_PKG_Maker.Font = New System.Drawing.Font("굴림", 9.5!)
        Me.CB_PKG_Maker.FormattingEnabled = True
        Me.CB_PKG_Maker.Location = New System.Drawing.Point(229, 191)
        Me.CB_PKG_Maker.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.CB_PKG_Maker.Name = "CB_PKG_Maker"
        Me.CB_PKG_Maker.Size = New System.Drawing.Size(103, 21)
        Me.CB_PKG_Maker.TabIndex = 27
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(9, 191)
        Me.Label9.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(220, 21)
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "PKG (Maker, Part No, 사용수량)"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CB_ModelDivision
        '
        Me.CB_ModelDivision.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_ModelDivision.Font = New System.Drawing.Font("굴림", 9.5!)
        Me.CB_ModelDivision.FormattingEnabled = True
        Me.CB_ModelDivision.Items.AddRange(New Object() {"개발", "양산"})
        Me.CB_ModelDivision.Location = New System.Drawing.Point(229, 32)
        Me.CB_ModelDivision.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.CB_ModelDivision.Name = "CB_ModelDivision"
        Me.CB_ModelDivision.Size = New System.Drawing.Size(103, 21)
        Me.CB_ModelDivision.TabIndex = 4
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(9, 32)
        Me.Label8.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(220, 21)
        Me.Label8.TabIndex = 3
        Me.Label8.Text = "구분"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'MARKING_DELETE
        '
        Me.MARKING_DELETE.Location = New System.Drawing.Point(814, 134)
        Me.MARKING_DELETE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.MARKING_DELETE.Name = "MARKING_DELETE"
        Me.MARKING_DELETE.Size = New System.Drawing.Size(75, 23)
        Me.MARKING_DELETE.TabIndex = 25
        Me.MARKING_DELETE.Text = "Delete"
        Me.MARKING_DELETE.UseVisualStyleBackColor = True
        '
        'FLOOR_PLAN_DELETE
        '
        Me.FLOOR_PLAN_DELETE.Location = New System.Drawing.Point(814, 111)
        Me.FLOOR_PLAN_DELETE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.FLOOR_PLAN_DELETE.Name = "FLOOR_PLAN_DELETE"
        Me.FLOOR_PLAN_DELETE.Size = New System.Drawing.Size(75, 23)
        Me.FLOOR_PLAN_DELETE.TabIndex = 19
        Me.FLOOR_PLAN_DELETE.Text = "Delete"
        Me.FLOOR_PLAN_DELETE.UseVisualStyleBackColor = True
        '
        'APP_SHEET_DELETE
        '
        Me.APP_SHEET_DELETE.Location = New System.Drawing.Point(814, 88)
        Me.APP_SHEET_DELETE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.APP_SHEET_DELETE.Name = "APP_SHEET_DELETE"
        Me.APP_SHEET_DELETE.Size = New System.Drawing.Size(75, 23)
        Me.APP_SHEET_DELETE.TabIndex = 13
        Me.APP_SHEET_DELETE.Text = "Delete"
        Me.APP_SHEET_DELETE.UseVisualStyleBackColor = True
        '
        'MARKING_DOWNLOAD
        '
        Me.MARKING_DOWNLOAD.Location = New System.Drawing.Point(739, 134)
        Me.MARKING_DOWNLOAD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.MARKING_DOWNLOAD.Name = "MARKING_DOWNLOAD"
        Me.MARKING_DOWNLOAD.Size = New System.Drawing.Size(75, 23)
        Me.MARKING_DOWNLOAD.TabIndex = 24
        Me.MARKING_DOWNLOAD.Text = "Download"
        Me.MARKING_DOWNLOAD.UseVisualStyleBackColor = True
        '
        'MARKING_FILE
        '
        Me.MARKING_FILE.BackColor = System.Drawing.SystemColors.Window
        Me.MARKING_FILE.Enabled = False
        Me.MARKING_FILE.Location = New System.Drawing.Point(332, 135)
        Me.MARKING_FILE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.MARKING_FILE.Name = "MARKING_FILE"
        Me.MARKING_FILE.Size = New System.Drawing.Size(329, 21)
        Me.MARKING_FILE.TabIndex = 22
        Me.MARKING_FILE.Text = "등록안됨"
        '
        'MARKING_UPLOAD
        '
        Me.MARKING_UPLOAD.Location = New System.Drawing.Point(664, 134)
        Me.MARKING_UPLOAD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.MARKING_UPLOAD.Name = "MARKING_UPLOAD"
        Me.MARKING_UPLOAD.Size = New System.Drawing.Size(75, 23)
        Me.MARKING_UPLOAD.TabIndex = 23
        Me.MARKING_UPLOAD.Text = "Upload"
        Me.MARKING_UPLOAD.UseVisualStyleBackColor = True
        '
        'MARKING_DATE
        '
        Me.MARKING_DATE.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.MARKING_DATE.Location = New System.Drawing.Point(229, 135)
        Me.MARKING_DATE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.MARKING_DATE.Name = "MARKING_DATE"
        Me.MARKING_DATE.Size = New System.Drawing.Size(103, 21)
        Me.MARKING_DATE.TabIndex = 21
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(9, 135)
        Me.Label7.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(220, 21)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "마킹기준"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FLOOR_PLAN_DOWNLOAD
        '
        Me.FLOOR_PLAN_DOWNLOAD.Location = New System.Drawing.Point(739, 111)
        Me.FLOOR_PLAN_DOWNLOAD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.FLOOR_PLAN_DOWNLOAD.Name = "FLOOR_PLAN_DOWNLOAD"
        Me.FLOOR_PLAN_DOWNLOAD.Size = New System.Drawing.Size(75, 23)
        Me.FLOOR_PLAN_DOWNLOAD.TabIndex = 18
        Me.FLOOR_PLAN_DOWNLOAD.Text = "Download"
        Me.FLOOR_PLAN_DOWNLOAD.UseVisualStyleBackColor = True
        '
        'FLOOR_PLAN_FILE
        '
        Me.FLOOR_PLAN_FILE.BackColor = System.Drawing.SystemColors.Window
        Me.FLOOR_PLAN_FILE.Enabled = False
        Me.FLOOR_PLAN_FILE.Location = New System.Drawing.Point(332, 112)
        Me.FLOOR_PLAN_FILE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.FLOOR_PLAN_FILE.Name = "FLOOR_PLAN_FILE"
        Me.FLOOR_PLAN_FILE.Size = New System.Drawing.Size(329, 21)
        Me.FLOOR_PLAN_FILE.TabIndex = 16
        Me.FLOOR_PLAN_FILE.Text = "등록안됨"
        '
        'FLOOR_PLAN_UPLOAD
        '
        Me.FLOOR_PLAN_UPLOAD.Location = New System.Drawing.Point(664, 111)
        Me.FLOOR_PLAN_UPLOAD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.FLOOR_PLAN_UPLOAD.Name = "FLOOR_PLAN_UPLOAD"
        Me.FLOOR_PLAN_UPLOAD.Size = New System.Drawing.Size(75, 23)
        Me.FLOOR_PLAN_UPLOAD.TabIndex = 17
        Me.FLOOR_PLAN_UPLOAD.Text = "Upload"
        Me.FLOOR_PLAN_UPLOAD.UseVisualStyleBackColor = True
        '
        'FLOOR_PLAN_DATE
        '
        Me.FLOOR_PLAN_DATE.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.FLOOR_PLAN_DATE.Location = New System.Drawing.Point(229, 112)
        Me.FLOOR_PLAN_DATE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.FLOOR_PLAN_DATE.Name = "FLOOR_PLAN_DATE"
        Me.FLOOR_PLAN_DATE.Size = New System.Drawing.Size(103, 21)
        Me.FLOOR_PLAN_DATE.TabIndex = 15
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(9, 112)
        Me.Label6.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(220, 21)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "도면"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'APP_SHEET_DOWNLOAD
        '
        Me.APP_SHEET_DOWNLOAD.Location = New System.Drawing.Point(739, 88)
        Me.APP_SHEET_DOWNLOAD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.APP_SHEET_DOWNLOAD.Name = "APP_SHEET_DOWNLOAD"
        Me.APP_SHEET_DOWNLOAD.Size = New System.Drawing.Size(75, 23)
        Me.APP_SHEET_DOWNLOAD.TabIndex = 12
        Me.APP_SHEET_DOWNLOAD.Text = "Download"
        Me.APP_SHEET_DOWNLOAD.UseVisualStyleBackColor = True
        '
        'APP_SHEET_FILE
        '
        Me.APP_SHEET_FILE.BackColor = System.Drawing.SystemColors.Window
        Me.APP_SHEET_FILE.Enabled = False
        Me.APP_SHEET_FILE.Location = New System.Drawing.Point(332, 89)
        Me.APP_SHEET_FILE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.APP_SHEET_FILE.Name = "APP_SHEET_FILE"
        Me.APP_SHEET_FILE.Size = New System.Drawing.Size(329, 21)
        Me.APP_SHEET_FILE.TabIndex = 10
        Me.APP_SHEET_FILE.Text = "등록안됨"
        '
        'APP_SHEET_UPLOAD
        '
        Me.APP_SHEET_UPLOAD.Location = New System.Drawing.Point(664, 88)
        Me.APP_SHEET_UPLOAD.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.APP_SHEET_UPLOAD.Name = "APP_SHEET_UPLOAD"
        Me.APP_SHEET_UPLOAD.Size = New System.Drawing.Size(75, 23)
        Me.APP_SHEET_UPLOAD.TabIndex = 11
        Me.APP_SHEET_UPLOAD.Text = "Upload"
        Me.APP_SHEET_UPLOAD.UseVisualStyleBackColor = True
        '
        'APP_SHEET_DATE
        '
        Me.APP_SHEET_DATE.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.APP_SHEET_DATE.Location = New System.Drawing.Point(229, 89)
        Me.APP_SHEET_DATE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.APP_SHEET_DATE.Name = "APP_SHEET_DATE"
        Me.APP_SHEET_DATE.Size = New System.Drawing.Size(103, 21)
        Me.APP_SHEET_DATE.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(9, 89)
        Me.Label5.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(220, 21)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "승인원"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'MODEL_NAME
        '
        Me.MODEL_NAME.BackColor = System.Drawing.SystemColors.Window
        Me.MODEL_NAME.Location = New System.Drawing.Point(332, 55)
        Me.MODEL_NAME.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.MODEL_NAME.Name = "MODEL_NAME"
        Me.MODEL_NAME.Size = New System.Drawing.Size(329, 21)
        Me.MODEL_NAME.TabIndex = 7
        '
        'MODEL_CODE
        '
        Me.MODEL_CODE.BackColor = System.Drawing.SystemColors.Window
        Me.MODEL_CODE.Enabled = False
        Me.MODEL_CODE.Location = New System.Drawing.Point(229, 55)
        Me.MODEL_CODE.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.MODEL_CODE.Name = "MODEL_CODE"
        Me.MODEL_CODE.Size = New System.Drawing.Size(103, 21)
        Me.MODEL_CODE.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(9, 55)
        Me.Label4.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(220, 21)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Model(Code/Name/PN)"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_Search, Me.BTN_NewModel, Me.BTN_NewCancel, Me.BTN_ModelSave})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1394, 25)
        Me.ToolStrip1.TabIndex = 0
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'BTN_Search
        '
        Me.BTN_Search.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.Search
        Me.BTN_Search.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_Search.Name = "BTN_Search"
        Me.BTN_Search.Size = New System.Drawing.Size(51, 22)
        Me.BTN_Search.Text = "조회"
        '
        'BTN_NewModel
        '
        Me.BTN_NewModel.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.plus
        Me.BTN_NewModel.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_NewModel.Name = "BTN_NewModel"
        Me.BTN_NewModel.Size = New System.Drawing.Size(75, 22)
        Me.BTN_NewModel.Text = "신규등록"
        '
        'BTN_NewCancel
        '
        Me.BTN_NewCancel.Enabled = False
        Me.BTN_NewCancel.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.cancel
        Me.BTN_NewCancel.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_NewCancel.Name = "BTN_NewCancel"
        Me.BTN_NewCancel.Size = New System.Drawing.Size(103, 22)
        Me.BTN_NewCancel.Text = "신규등록 취소"
        '
        'BTN_ModelSave
        '
        Me.BTN_ModelSave.Enabled = False
        Me.BTN_ModelSave.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.save
        Me.BTN_ModelSave.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_ModelSave.Name = "BTN_ModelSave"
        Me.BTN_ModelSave.Size = New System.Drawing.Size(51, 22)
        Me.BTN_ModelSave.Text = "저장"
        '
        'GRID_MENU
        '
        Me.GRID_MENU.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_ADD, Me.BTN_DELETE, Me.ToolStripSeparator1, Me.BTN_Save})
        Me.GRID_MENU.Name = "GRID_MENU"
        Me.GRID_MENU.Size = New System.Drawing.Size(123, 76)
        '
        'BTN_ADD
        '
        Me.BTN_ADD.Enabled = False
        Me.BTN_ADD.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.plus
        Me.BTN_ADD.Name = "BTN_ADD"
        Me.BTN_ADD.Size = New System.Drawing.Size(122, 22)
        Me.BTN_ADD.Text = "추가"
        '
        'BTN_DELETE
        '
        Me.BTN_DELETE.Name = "BTN_DELETE"
        Me.BTN_DELETE.Size = New System.Drawing.Size(122, 22)
        Me.BTN_DELETE.Text = "사용안함"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(119, 6)
        '
        'BTN_Save
        '
        Me.BTN_Save.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.save
        Me.BTN_Save.Name = "BTN_Save"
        Me.BTN_Save.Size = New System.Drawing.Size(122, 22)
        Me.BTN_Save.Text = "저장"
        '
        'GRID_MENU2
        '
        Me.GRID_MENU2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_ROW_ADD, Me.BTN_ROW_DELETE, Me.BTN_COL_ADD, Me.BTN_COL_DELETE})
        Me.GRID_MENU2.Name = "GRID_MENU"
        Me.GRID_MENU2.Size = New System.Drawing.Size(115, 92)
        '
        'BTN_ROW_ADD
        '
        Me.BTN_ROW_ADD.Name = "BTN_ROW_ADD"
        Me.BTN_ROW_ADD.Size = New System.Drawing.Size(114, 22)
        Me.BTN_ROW_ADD.Text = "행 추가"
        '
        'BTN_ROW_DELETE
        '
        Me.BTN_ROW_DELETE.Name = "BTN_ROW_DELETE"
        Me.BTN_ROW_DELETE.Size = New System.Drawing.Size(114, 22)
        Me.BTN_ROW_DELETE.Text = "행 삭제"
        '
        'BTN_COL_ADD
        '
        Me.BTN_COL_ADD.Name = "BTN_COL_ADD"
        Me.BTN_COL_ADD.Size = New System.Drawing.Size(114, 22)
        Me.BTN_COL_ADD.Text = "열 추가"
        '
        'BTN_COL_DELETE
        '
        Me.BTN_COL_DELETE.Name = "BTN_COL_DELETE"
        Me.BTN_COL_DELETE.Size = New System.Drawing.Size(114, 22)
        Me.BTN_COL_DELETE.Text = "열 삭제"
        '
        'BluModelManager
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1394, 759)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "BluModelManager"
        Me.Text = "모델 등록 / 관리"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.GRID_ModelList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PACKING_QTY, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TH_QTY, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PKG_QTY, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.GRID_MENU.ResumeLayout(False)
        Me.GRID_MENU2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GRID_ModelList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents CB_Model_Division As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TB_Model_Name As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents BTN_Search As System.Windows.Forms.ToolStripButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents RB_Model_Use As System.Windows.Forms.RadioButton
    Friend WithEvents RB_Model_Not_Use As System.Windows.Forms.RadioButton
    Friend WithEvents RB_Model_ALL As System.Windows.Forms.RadioButton
    Friend WithEvents GRID_MENU As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents BTN_ADD As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BTN_DELETE As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BTN_Save As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents MODEL_NAME As System.Windows.Forms.TextBox
    Friend WithEvents MODEL_CODE As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents MARKING_DOWNLOAD As System.Windows.Forms.Button
    Friend WithEvents MARKING_FILE As System.Windows.Forms.TextBox
    Friend WithEvents MARKING_UPLOAD As System.Windows.Forms.Button
    Friend WithEvents MARKING_DATE As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents FLOOR_PLAN_DOWNLOAD As System.Windows.Forms.Button
    Friend WithEvents FLOOR_PLAN_FILE As System.Windows.Forms.TextBox
    Friend WithEvents FLOOR_PLAN_UPLOAD As System.Windows.Forms.Button
    Friend WithEvents FLOOR_PLAN_DATE As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents APP_SHEET_DOWNLOAD As System.Windows.Forms.Button
    Friend WithEvents APP_SHEET_FILE As System.Windows.Forms.TextBox
    Friend WithEvents APP_SHEET_UPLOAD As System.Windows.Forms.Button
    Friend WithEvents APP_SHEET_DATE As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents APP_SHEET_DELETE As System.Windows.Forms.Button
    Friend WithEvents MARKING_DELETE As System.Windows.Forms.Button
    Friend WithEvents FLOOR_PLAN_DELETE As System.Windows.Forms.Button
    Friend WithEvents CB_ModelDivision As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents BTN_NewModel As System.Windows.Forms.ToolStripButton
    Friend WithEvents BTN_NewCancel As System.Windows.Forms.ToolStripButton
    Friend WithEvents BTN_ModelSave As System.Windows.Forms.ToolStripButton
    Friend WithEvents PKG_PARTNO As System.Windows.Forms.TextBox
    Friend WithEvents CB_PKG_Maker As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents CB_RankOrder As System.Windows.Forms.ComboBox
    Friend WithEvents TH_PARTNO As System.Windows.Forms.TextBox
    Friend WithEvents CB_TH_Maker As System.Windows.Forms.ComboBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents GRID_MENU2 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents BTN_ROW_ADD As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BTN_ROW_DELETE As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BTN_COL_ADD As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BTN_COL_DELETE As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MODEL_NOTE As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents RB_NOT_USE As System.Windows.Forms.RadioButton
    Friend WithEvents RB_USE As System.Windows.Forms.RadioButton
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents PCB_PARTNO As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TH_QTY As System.Windows.Forms.NumericUpDown
    Friend WithEvents PKG_QTY As System.Windows.Forms.NumericUpDown
    Friend WithEvents INTENSITY_RANK As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents CHROMATICITY_RANK As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents PACKING_QTY As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents MODEL_PART_NO As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents OFP_DELETE As System.Windows.Forms.Button
    Friend WithEvents OFP_DOWNLOAD As System.Windows.Forms.Button
    Friend WithEvents OQC_FLOOR_PLAN_FILE As System.Windows.Forms.TextBox
    Friend WithEvents OFP_UPLOAD As System.Windows.Forms.Button
    Friend WithEvents OQC_FLOOR_PLAN_DATE As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label20 As System.Windows.Forms.Label
End Class

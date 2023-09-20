<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_OMS_WIP_DATA
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_OMS_WIP_DATA))
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.BTN_Search = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_Data_Output = New System.Windows.Forms.ToolStripButton()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        Me.BTN_Table_Show = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.GRID_Data_List = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.TB_C_Area_ID = New System.Windows.Forms.TextBox()
        Me.TB_M_Area_ID = New System.Windows.Forms.TextBox()
        Me.TB_C_Area_Flag = New System.Windows.Forms.TextBox()
        Me.TB_M_Area_Flag = New System.Windows.Forms.TextBox()
        Me.TB_C_Step_Desc = New System.Windows.Forms.TextBox()
        Me.TB_M_Step_Desc = New System.Windows.Forms.TextBox()
        Me.TB_C_Step_SEQ = New System.Windows.Forms.TextBox()
        Me.TB_M_Step_SEQ = New System.Windows.Forms.TextBox()
        Me.TB_C_Step_ID = New System.Windows.Forms.TextBox()
        Me.TB_M_Step_ID = New System.Windows.Forms.TextBox()
        Me.TB_C_Process_ID = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TB_M_Process_ID = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TB_Inkless = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.TB_Wafer_ID = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.TB_Bonus_Qty = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.TB_Loss_QTY = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.TB_Other = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.TB_NCQ_Code = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.TB_NCT_Code = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.TB_NCA_Code = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.TB_NCF_Code = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.TB_Hold_Datetime = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.TB_Hold_Code = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TB_Start_Tag = New System.Windows.Forms.TextBox()
        Me.TB_Hold_Flag = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TB_Company_Code = New System.Windows.Forms.TextBox()
        Me.TB_Item_Flag = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.GRID_Lot_List = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cms_Menu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.btn_RowDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_Data_Output2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.GRID_Data_List, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.GRID_Lot_List, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cms_Menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_Search, Me.ToolStripSeparator1, Me.BTN_Data_Output, Me.Form_CLose, Me.BTN_Table_Show, Me.ToolStripButton1})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1664, 25)
        Me.ToolStrip1.TabIndex = 7
        Me.ToolStrip1.Text = "ToolStrip1"
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
        'BTN_Data_Output
        '
        Me.BTN_Data_Output.Image = Global.Repair_System.My.Resources.Resources.TEST_FOLDER
        Me.BTN_Data_Output.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_Data_Output.Name = "BTN_Data_Output"
        Me.BTN_Data_Output.Size = New System.Drawing.Size(115, 22)
        Me.BTN_Data_Output.Text = "재공데이터 작성"
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
        'BTN_Table_Show
        '
        Me.BTN_Table_Show.Image = Global.Repair_System.My.Resources.Resources.folder_black
        Me.BTN_Table_Show.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_Table_Show.Name = "BTN_Table_Show"
        Me.BTN_Table_Show.Size = New System.Drawing.Size(83, 22)
        Me.BTN_Table_Show.Text = "Table 보기"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(73, 22)
        Me.ToolStripButton1.Text = "'-1' 변환"
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.GRID_Data_List)
        Me.SplitContainer1.Panel1Collapsed = True
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1664, 748)
        Me.SplitContainer1.SplitterDistance = 1229
        Me.SplitContainer1.TabIndex = 8
        '
        'GRID_Data_List
        '
        Me.GRID_Data_List.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.GRID_Data_List.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GRID_Data_List.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.GRID_Data_List.Location = New System.Drawing.Point(0, 0)
        Me.GRID_Data_List.Name = "GRID_Data_List"
        Me.GRID_Data_List.Rows.Count = 2
        Me.GRID_Data_List.Rows.DefaultSize = 20
        Me.GRID_Data_List.Size = New System.Drawing.Size(1229, 100)
        Me.GRID_Data_List.StyleInfo = resources.GetString("GRID_Data_List.StyleInfo")
        Me.GRID_Data_List.TabIndex = 55
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
        Me.SplitContainer2.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label22)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Panel2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label25)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GRID_Lot_List)
        Me.SplitContainer2.Size = New System.Drawing.Size(1664, 748)
        Me.SplitContainer2.SplitterDistance = 127
        Me.SplitContainer2.TabIndex = 0
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.ForeColor = System.Drawing.Color.Black
        Me.Label22.Location = New System.Drawing.Point(3, 111)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(111, 12)
        Me.Label22.TabIndex = 34
        Me.Label22.Text = "총 모듈 수량 : 0 EA"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.TB_C_Area_ID)
        Me.Panel2.Controls.Add(Me.TB_M_Area_ID)
        Me.Panel2.Controls.Add(Me.TB_C_Area_Flag)
        Me.Panel2.Controls.Add(Me.TB_M_Area_Flag)
        Me.Panel2.Controls.Add(Me.TB_C_Step_Desc)
        Me.Panel2.Controls.Add(Me.TB_M_Step_Desc)
        Me.Panel2.Controls.Add(Me.TB_C_Step_SEQ)
        Me.Panel2.Controls.Add(Me.TB_M_Step_SEQ)
        Me.Panel2.Controls.Add(Me.TB_C_Step_ID)
        Me.Panel2.Controls.Add(Me.TB_M_Step_ID)
        Me.Panel2.Controls.Add(Me.TB_C_Process_ID)
        Me.Panel2.Controls.Add(Me.Label12)
        Me.Panel2.Controls.Add(Me.Label11)
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.TB_M_Process_ID)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Location = New System.Drawing.Point(1008, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(618, 120)
        Me.Panel2.TabIndex = 18
        '
        'TB_C_Area_ID
        '
        Me.TB_C_Area_ID.BackColor = System.Drawing.SystemColors.Window
        Me.TB_C_Area_ID.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_C_Area_ID.Location = New System.Drawing.Point(516, 61)
        Me.TB_C_Area_ID.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_C_Area_ID.Name = "TB_C_Area_ID"
        Me.TB_C_Area_ID.Size = New System.Drawing.Size(86, 21)
        Me.TB_C_Area_ID.TabIndex = 33
        Me.TB_C_Area_ID.Text = "JMCS"
        '
        'TB_M_Area_ID
        '
        Me.TB_M_Area_ID.BackColor = System.Drawing.SystemColors.Window
        Me.TB_M_Area_ID.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_M_Area_ID.Location = New System.Drawing.Point(516, 37)
        Me.TB_M_Area_ID.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_M_Area_ID.Name = "TB_M_Area_ID"
        Me.TB_M_Area_ID.Size = New System.Drawing.Size(86, 21)
        Me.TB_M_Area_ID.TabIndex = 32
        Me.TB_M_Area_ID.Text = "JMDL"
        '
        'TB_C_Area_Flag
        '
        Me.TB_C_Area_Flag.BackColor = System.Drawing.SystemColors.Window
        Me.TB_C_Area_Flag.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_C_Area_Flag.Location = New System.Drawing.Point(430, 61)
        Me.TB_C_Area_Flag.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_C_Area_Flag.Name = "TB_C_Area_Flag"
        Me.TB_C_Area_Flag.Size = New System.Drawing.Size(86, 21)
        Me.TB_C_Area_Flag.TabIndex = 31
        Me.TB_C_Area_Flag.Text = "ASSY"
        '
        'TB_M_Area_Flag
        '
        Me.TB_M_Area_Flag.BackColor = System.Drawing.SystemColors.Window
        Me.TB_M_Area_Flag.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_M_Area_Flag.Location = New System.Drawing.Point(430, 37)
        Me.TB_M_Area_Flag.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_M_Area_Flag.Name = "TB_M_Area_Flag"
        Me.TB_M_Area_Flag.Size = New System.Drawing.Size(86, 21)
        Me.TB_M_Area_Flag.TabIndex = 30
        Me.TB_M_Area_Flag.Text = "TEST"
        '
        'TB_C_Step_Desc
        '
        Me.TB_C_Step_Desc.BackColor = System.Drawing.SystemColors.Window
        Me.TB_C_Step_Desc.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_C_Step_Desc.Location = New System.Drawing.Point(344, 61)
        Me.TB_C_Step_Desc.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_C_Step_Desc.Name = "TB_C_Step_Desc"
        Me.TB_C_Step_Desc.Size = New System.Drawing.Size(86, 21)
        Me.TB_C_Step_Desc.TabIndex = 29
        Me.TB_C_Step_Desc.Text = "HOME STOCK"
        '
        'TB_M_Step_Desc
        '
        Me.TB_M_Step_Desc.BackColor = System.Drawing.SystemColors.Window
        Me.TB_M_Step_Desc.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_M_Step_Desc.Location = New System.Drawing.Point(344, 37)
        Me.TB_M_Step_Desc.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_M_Step_Desc.Name = "TB_M_Step_Desc"
        Me.TB_M_Step_Desc.Size = New System.Drawing.Size(86, 21)
        Me.TB_M_Step_Desc.TabIndex = 28
        Me.TB_M_Step_Desc.Text = "REPAIR"
        '
        'TB_C_Step_SEQ
        '
        Me.TB_C_Step_SEQ.BackColor = System.Drawing.SystemColors.Window
        Me.TB_C_Step_SEQ.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_C_Step_SEQ.Location = New System.Drawing.Point(258, 61)
        Me.TB_C_Step_SEQ.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_C_Step_SEQ.Name = "TB_C_Step_SEQ"
        Me.TB_C_Step_SEQ.Size = New System.Drawing.Size(86, 21)
        Me.TB_C_Step_SEQ.TabIndex = 27
        Me.TB_C_Step_SEQ.Text = "A0001"
        '
        'TB_M_Step_SEQ
        '
        Me.TB_M_Step_SEQ.BackColor = System.Drawing.SystemColors.Window
        Me.TB_M_Step_SEQ.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_M_Step_SEQ.Location = New System.Drawing.Point(258, 37)
        Me.TB_M_Step_SEQ.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_M_Step_SEQ.Name = "TB_M_Step_SEQ"
        Me.TB_M_Step_SEQ.Size = New System.Drawing.Size(86, 21)
        Me.TB_M_Step_SEQ.TabIndex = 26
        Me.TB_M_Step_SEQ.Text = "M0500"
        '
        'TB_C_Step_ID
        '
        Me.TB_C_Step_ID.BackColor = System.Drawing.SystemColors.Window
        Me.TB_C_Step_ID.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_C_Step_ID.Location = New System.Drawing.Point(172, 61)
        Me.TB_C_Step_ID.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_C_Step_ID.Name = "TB_C_Step_ID"
        Me.TB_C_Step_ID.Size = New System.Drawing.Size(86, 21)
        Me.TB_C_Step_ID.TabIndex = 25
        Me.TB_C_Step_ID.Text = "0002"
        '
        'TB_M_Step_ID
        '
        Me.TB_M_Step_ID.BackColor = System.Drawing.SystemColors.Window
        Me.TB_M_Step_ID.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_M_Step_ID.Location = New System.Drawing.Point(172, 37)
        Me.TB_M_Step_ID.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_M_Step_ID.Name = "TB_M_Step_ID"
        Me.TB_M_Step_ID.Size = New System.Drawing.Size(86, 21)
        Me.TB_M_Step_ID.TabIndex = 24
        Me.TB_M_Step_ID.Text = "0007"
        '
        'TB_C_Process_ID
        '
        Me.TB_C_Process_ID.BackColor = System.Drawing.SystemColors.Window
        Me.TB_C_Process_ID.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_C_Process_ID.Location = New System.Drawing.Point(86, 61)
        Me.TB_C_Process_ID.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_C_Process_ID.Name = "TB_C_Process_ID"
        Me.TB_C_Process_ID.Size = New System.Drawing.Size(86, 21)
        Me.TB_C_Process_ID.TabIndex = 23
        Me.TB_C_Process_ID.Text = "BASIC"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label12.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.White
        Me.Label12.Location = New System.Drawing.Point(516, 14)
        Me.Label12.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(86, 21)
        Me.Label12.TabIndex = 22
        Me.Label12.Text = "Area ID"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label11.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(430, 14)
        Me.Label11.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(86, 21)
        Me.Label11.TabIndex = 21
        Me.Label11.Text = "Area Flag"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label10.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(344, 14)
        Me.Label10.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(86, 21)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "Step Desc"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(258, 14)
        Me.Label9.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(86, 21)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "Step SEQ No"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(172, 14)
        Me.Label8.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(86, 21)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Step ID"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_M_Process_ID
        '
        Me.TB_M_Process_ID.BackColor = System.Drawing.SystemColors.Window
        Me.TB_M_Process_ID.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_M_Process_ID.Location = New System.Drawing.Point(86, 37)
        Me.TB_M_Process_ID.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_M_Process_ID.Name = "TB_M_Process_ID"
        Me.TB_M_Process_ID.Size = New System.Drawing.Size(86, 21)
        Me.TB_M_Process_ID.TabIndex = 17
        Me.TB_M_Process_ID.Text = "JMDL"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(86, 14)
        Me.Label7.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(86, 21)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Process ID"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(3, 61)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 21)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Component"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(3, 37)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 21)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Module"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.TB_Inkless)
        Me.Panel1.Controls.Add(Me.Label21)
        Me.Panel1.Controls.Add(Me.TB_Wafer_ID)
        Me.Panel1.Controls.Add(Me.Label23)
        Me.Panel1.Controls.Add(Me.TB_Bonus_Qty)
        Me.Panel1.Controls.Add(Me.Label24)
        Me.Panel1.Controls.Add(Me.TB_Loss_QTY)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.TB_Other)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.TB_NCQ_Code)
        Me.Panel1.Controls.Add(Me.Label19)
        Me.Panel1.Controls.Add(Me.TB_NCT_Code)
        Me.Panel1.Controls.Add(Me.Label20)
        Me.Panel1.Controls.Add(Me.TB_NCA_Code)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.TB_NCF_Code)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.TB_Hold_Datetime)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.TB_Hold_Code)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.TB_Start_Tag)
        Me.Panel1.Controls.Add(Me.TB_Hold_Flag)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.TB_Company_Code)
        Me.Panel1.Controls.Add(Me.TB_Item_Flag)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Location = New System.Drawing.Point(177, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(830, 120)
        Me.Panel1.TabIndex = 17
        '
        'TB_Inkless
        '
        Me.TB_Inkless.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Inkless.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Inkless.Location = New System.Drawing.Point(747, 60)
        Me.TB_Inkless.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Inkless.Name = "TB_Inkless"
        Me.TB_Inkless.Size = New System.Drawing.Size(68, 21)
        Me.TB_Inkless.TabIndex = 36
        '
        'Label21
        '
        Me.Label21.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label21.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label21.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.White
        Me.Label21.Location = New System.Drawing.Point(614, 60)
        Me.Label21.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(133, 21)
        Me.Label21.TabIndex = 35
        Me.Label21.Text = "Inkless"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Wafer_ID
        '
        Me.TB_Wafer_ID.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Wafer_ID.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Wafer_ID.Location = New System.Drawing.Point(747, 37)
        Me.TB_Wafer_ID.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Wafer_ID.Name = "TB_Wafer_ID"
        Me.TB_Wafer_ID.Size = New System.Drawing.Size(68, 21)
        Me.TB_Wafer_ID.TabIndex = 36
        '
        'Label23
        '
        Me.Label23.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label23.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label23.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label23.ForeColor = System.Drawing.Color.White
        Me.Label23.Location = New System.Drawing.Point(614, 37)
        Me.Label23.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(133, 21)
        Me.Label23.TabIndex = 35
        Me.Label23.Text = "Wafer ID"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Bonus_Qty
        '
        Me.TB_Bonus_Qty.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Bonus_Qty.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Bonus_Qty.Location = New System.Drawing.Point(747, 14)
        Me.TB_Bonus_Qty.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Bonus_Qty.Name = "TB_Bonus_Qty"
        Me.TB_Bonus_Qty.Size = New System.Drawing.Size(68, 21)
        Me.TB_Bonus_Qty.TabIndex = 34
        Me.TB_Bonus_Qty.Text = "0"
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label24.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label24.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.White
        Me.Label24.Location = New System.Drawing.Point(614, 14)
        Me.Label24.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(133, 21)
        Me.Label24.TabIndex = 33
        Me.Label24.Text = "Bonus Qty"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Loss_QTY
        '
        Me.TB_Loss_QTY.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Loss_QTY.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Loss_QTY.Location = New System.Drawing.Point(546, 83)
        Me.TB_Loss_QTY.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Loss_QTY.Name = "TB_Loss_QTY"
        Me.TB_Loss_QTY.Size = New System.Drawing.Size(68, 21)
        Me.TB_Loss_QTY.TabIndex = 32
        Me.TB_Loss_QTY.Text = "0"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label17.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.White
        Me.Label17.Location = New System.Drawing.Point(413, 83)
        Me.Label17.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(133, 21)
        Me.Label17.TabIndex = 31
        Me.Label17.Text = "Loss Qty"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Other
        '
        Me.TB_Other.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Other.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Other.Location = New System.Drawing.Point(546, 60)
        Me.TB_Other.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Other.Name = "TB_Other"
        Me.TB_Other.Size = New System.Drawing.Size(68, 21)
        Me.TB_Other.TabIndex = 30
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label18.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label18.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.White
        Me.Label18.Location = New System.Drawing.Point(413, 60)
        Me.Label18.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(133, 21)
        Me.Label18.TabIndex = 29
        Me.Label18.Text = "Other"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_NCQ_Code
        '
        Me.TB_NCQ_Code.BackColor = System.Drawing.SystemColors.Window
        Me.TB_NCQ_Code.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_NCQ_Code.Location = New System.Drawing.Point(546, 37)
        Me.TB_NCQ_Code.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_NCQ_Code.Name = "TB_NCQ_Code"
        Me.TB_NCQ_Code.Size = New System.Drawing.Size(68, 21)
        Me.TB_NCQ_Code.TabIndex = 28
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label19.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.White
        Me.Label19.Location = New System.Drawing.Point(413, 37)
        Me.Label19.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(133, 21)
        Me.Label19.TabIndex = 27
        Me.Label19.Text = "NCQ Code"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_NCT_Code
        '
        Me.TB_NCT_Code.BackColor = System.Drawing.SystemColors.Window
        Me.TB_NCT_Code.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_NCT_Code.Location = New System.Drawing.Point(546, 14)
        Me.TB_NCT_Code.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_NCT_Code.Name = "TB_NCT_Code"
        Me.TB_NCT_Code.Size = New System.Drawing.Size(68, 21)
        Me.TB_NCT_Code.TabIndex = 26
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label20.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label20.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.White
        Me.Label20.Location = New System.Drawing.Point(413, 14)
        Me.Label20.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(133, 21)
        Me.Label20.TabIndex = 25
        Me.Label20.Text = "NCT Code"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_NCA_Code
        '
        Me.TB_NCA_Code.BackColor = System.Drawing.SystemColors.Window
        Me.TB_NCA_Code.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_NCA_Code.Location = New System.Drawing.Point(345, 83)
        Me.TB_NCA_Code.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_NCA_Code.Name = "TB_NCA_Code"
        Me.TB_NCA_Code.Size = New System.Drawing.Size(68, 21)
        Me.TB_NCA_Code.TabIndex = 24
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label16.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.White
        Me.Label16.Location = New System.Drawing.Point(212, 83)
        Me.Label16.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(133, 21)
        Me.Label16.TabIndex = 23
        Me.Label16.Text = "NCA Code"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_NCF_Code
        '
        Me.TB_NCF_Code.BackColor = System.Drawing.SystemColors.Window
        Me.TB_NCF_Code.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_NCF_Code.Location = New System.Drawing.Point(345, 60)
        Me.TB_NCF_Code.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_NCF_Code.Name = "TB_NCF_Code"
        Me.TB_NCF_Code.Size = New System.Drawing.Size(68, 21)
        Me.TB_NCF_Code.TabIndex = 22
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label15.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(212, 60)
        Me.Label15.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(133, 21)
        Me.Label15.TabIndex = 21
        Me.Label15.Text = "NCF Code"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Hold_Datetime
        '
        Me.TB_Hold_Datetime.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Hold_Datetime.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Hold_Datetime.Location = New System.Drawing.Point(345, 37)
        Me.TB_Hold_Datetime.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Hold_Datetime.Name = "TB_Hold_Datetime"
        Me.TB_Hold_Datetime.Size = New System.Drawing.Size(68, 21)
        Me.TB_Hold_Datetime.TabIndex = 20
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label14.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(212, 37)
        Me.Label14.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(133, 21)
        Me.Label14.TabIndex = 19
        Me.Label14.Text = "Hold DateTime"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Hold_Code
        '
        Me.TB_Hold_Code.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Hold_Code.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Hold_Code.Location = New System.Drawing.Point(345, 14)
        Me.TB_Hold_Code.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Hold_Code.Name = "TB_Hold_Code"
        Me.TB_Hold_Code.Size = New System.Drawing.Size(68, 21)
        Me.TB_Hold_Code.TabIndex = 18
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label13.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.White
        Me.Label13.Location = New System.Drawing.Point(212, 14)
        Me.Label13.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(133, 21)
        Me.Label13.TabIndex = 17
        Me.Label13.Text = "Hold Code"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(11, 14)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(133, 21)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Start Tag"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Start_Tag
        '
        Me.TB_Start_Tag.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Start_Tag.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Start_Tag.Location = New System.Drawing.Point(144, 14)
        Me.TB_Start_Tag.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Start_Tag.Name = "TB_Start_Tag"
        Me.TB_Start_Tag.Size = New System.Drawing.Size(68, 21)
        Me.TB_Start_Tag.TabIndex = 10
        Me.TB_Start_Tag.Text = "LS1"
        '
        'TB_Hold_Flag
        '
        Me.TB_Hold_Flag.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Hold_Flag.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Hold_Flag.Location = New System.Drawing.Point(144, 83)
        Me.TB_Hold_Flag.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Hold_Flag.Name = "TB_Hold_Flag"
        Me.TB_Hold_Flag.Size = New System.Drawing.Size(68, 21)
        Me.TB_Hold_Flag.TabIndex = 16
        Me.TB_Hold_Flag.Text = "R"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(11, 37)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(133, 21)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Company Code"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(11, 83)
        Me.Label4.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(133, 21)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Hold Flag"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Company_Code
        '
        Me.TB_Company_Code.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Company_Code.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Company_Code.Location = New System.Drawing.Point(144, 37)
        Me.TB_Company_Code.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Company_Code.Name = "TB_Company_Code"
        Me.TB_Company_Code.Size = New System.Drawing.Size(68, 21)
        Me.TB_Company_Code.TabIndex = 12
        Me.TB_Company_Code.Text = "UJ"
        '
        'TB_Item_Flag
        '
        Me.TB_Item_Flag.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Item_Flag.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Item_Flag.Location = New System.Drawing.Point(144, 60)
        Me.TB_Item_Flag.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Item_Flag.Name = "TB_Item_Flag"
        Me.TB_Item_Flag.Size = New System.Drawing.Size(68, 21)
        Me.TB_Item_Flag.TabIndex = 14
        Me.TB_Item_Flag.Text = "C"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(11, 60)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(133, 21)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Wafer Chip Flag"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.ForeColor = System.Drawing.Color.Black
        Me.Label25.Location = New System.Drawing.Point(3, 94)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(105, 12)
        Me.Label25.TabIndex = 35
        Me.Label25.Text = "총 Lot 수량 : 0 Lot"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GRID_Lot_List
        '
        Me.GRID_Lot_List.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.GRID_Lot_List.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GRID_Lot_List.Location = New System.Drawing.Point(0, 0)
        Me.GRID_Lot_List.Name = "GRID_Lot_List"
        Me.GRID_Lot_List.Rows.Count = 2
        Me.GRID_Lot_List.Rows.DefaultSize = 20
        Me.GRID_Lot_List.Size = New System.Drawing.Size(1664, 617)
        Me.GRID_Lot_List.StyleInfo = resources.GetString("GRID_Lot_List.StyleInfo")
        Me.GRID_Lot_List.TabIndex = 56
        '
        'cms_Menu
        '
        Me.cms_Menu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_RowDelete, Me.ToolStripSeparator4, Me.BTN_Data_Output2})
        Me.cms_Menu.Name = "GRID_MENU"
        Me.cms_Menu.Size = New System.Drawing.Size(163, 54)
        '
        'btn_RowDelete
        '
        Me.btn_RowDelete.Image = Global.Repair_System.My.Resources.Resources.minus_blue
        Me.btn_RowDelete.Name = "btn_RowDelete"
        Me.btn_RowDelete.Size = New System.Drawing.Size(162, 22)
        Me.btn_RowDelete.Text = "행 삭제"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(159, 6)
        '
        'BTN_Data_Output2
        '
        Me.BTN_Data_Output2.Image = Global.Repair_System.My.Resources.Resources.TEST_FOLDER
        Me.BTN_Data_Output2.Name = "BTN_Data_Output2"
        Me.BTN_Data_Output2.Size = New System.Drawing.Size(162, 22)
        Me.BTN_Data_Output2.Text = "재공데이터 작성"
        '
        'frm_OMS_WIP_DATA
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1664, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_OMS_WIP_DATA"
        Me.Text = "OMS 재공데이터 생성"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.GRID_Data_List, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.GRID_Lot_List, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cms_Menu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents BTN_Search As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BTN_Data_Output As System.Windows.Forms.ToolStripButton
    Friend WithEvents Form_CLose As System.Windows.Forms.ToolStripButton
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GRID_Data_List As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents BTN_Table_Show As System.Windows.Forms.ToolStripButton
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GRID_Lot_List As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents TB_Hold_Flag As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TB_Item_Flag As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TB_Company_Code As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TB_Start_Tag As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TB_C_Area_ID As System.Windows.Forms.TextBox
    Friend WithEvents TB_M_Area_ID As System.Windows.Forms.TextBox
    Friend WithEvents TB_C_Area_Flag As System.Windows.Forms.TextBox
    Friend WithEvents TB_M_Area_Flag As System.Windows.Forms.TextBox
    Friend WithEvents TB_C_Step_Desc As System.Windows.Forms.TextBox
    Friend WithEvents TB_M_Step_Desc As System.Windows.Forms.TextBox
    Friend WithEvents TB_C_Step_SEQ As System.Windows.Forms.TextBox
    Friend WithEvents TB_M_Step_SEQ As System.Windows.Forms.TextBox
    Friend WithEvents TB_C_Step_ID As System.Windows.Forms.TextBox
    Friend WithEvents TB_M_Step_ID As System.Windows.Forms.TextBox
    Friend WithEvents TB_C_Process_ID As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TB_M_Process_ID As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TB_NCA_Code As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TB_NCF_Code As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TB_Hold_Datetime As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TB_Hold_Code As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TB_Loss_QTY As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TB_Other As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TB_NCQ_Code As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TB_NCT_Code As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents TB_Wafer_ID As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents TB_Bonus_Qty As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents TB_Inkless As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents cms_Menu As ContextMenuStrip
    Friend WithEvents btn_RowDelete As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents BTN_Data_Output2 As ToolStripMenuItem
    Friend WithEvents Label22 As Label
    Friend WithEvents Label25 As Label
End Class

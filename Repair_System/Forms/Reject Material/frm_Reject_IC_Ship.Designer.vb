<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_Reject_IC_Ship
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Reject_IC_Ship))
        Me.btn_RowDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.btn_Search = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_New_Reject = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Save = New System.Windows.Forms.ToolStripButton()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        Me.btn_PMIC_Print = New System.Windows.Forms.ToolStripButton()
        Me.grid_RejectList_Detail = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.date_Reject_End = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.date_Reject_Start = New System.Windows.Forms.DateTimePicker()
        Me.rb_Date = New System.Windows.Forms.RadioButton()
        Me.rb_Area = New System.Windows.Forms.RadioButton()
        Me.DateTimePicker3 = New System.Windows.Forms.DateTimePicker()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btn_InputCheck = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tb_Max_No = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tb_Min_No = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.tb_S_YJ_No = New System.Windows.Forms.TextBox()
        Me.tb_S_Lot_No = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.grid_RejectList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.grid_PartList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cms_Menu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ToolStrip1.SuspendLayout()
        CType(Me.grid_RejectList_Detail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grid_RejectList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grid_PartList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cms_Menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_RowDelete
        '
        Me.btn_RowDelete.Image = Global.Repair_System.My.Resources.Resources.minus_blue
        Me.btn_RowDelete.Name = "btn_RowDelete"
        Me.btn_RowDelete.Size = New System.Drawing.Size(119, 22)
        Me.btn_RowDelete.Text = "Lot 삭제"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Search, Me.ToolStripSeparator1, Me.btn_New_Reject, Me.ToolStripSeparator2, Me.btn_Save, Me.Form_CLose, Me.btn_PMIC_Print})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1264, 25)
        Me.ToolStrip1.TabIndex = 9
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'btn_Search
        '
        Me.btn_Search.Image = Global.Repair_System.My.Resources.Resources.Search_121
        Me.btn_Search.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_Search.Name = "btn_Search"
        Me.btn_Search.Size = New System.Drawing.Size(51, 22)
        Me.btn_Search.Text = "조회"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'btn_New_Reject
        '
        Me.btn_New_Reject.Image = Global.Repair_System.My.Resources.Resources.TEST_FOLDER
        Me.btn_New_Reject.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_New_Reject.Name = "btn_New_Reject"
        Me.btn_New_Reject.Size = New System.Drawing.Size(75, 22)
        Me.btn_New_Reject.Text = "신규등록"
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
        'btn_PMIC_Print
        '
        Me.btn_PMIC_Print.Enabled = False
        Me.btn_PMIC_Print.Image = Global.Repair_System.My.Resources.Resources.print
        Me.btn_PMIC_Print.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_PMIC_Print.Name = "btn_PMIC_Print"
        Me.btn_PMIC_Print.Size = New System.Drawing.Size(103, 22)
        Me.btn_PMIC_Print.Text = "반납전표 발행"
        '
        'grid_RejectList_Detail
        '
        Me.grid_RejectList_Detail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grid_RejectList_Detail.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_RejectList_Detail.Location = New System.Drawing.Point(9, 177)
        Me.grid_RejectList_Detail.Name = "grid_RejectList_Detail"
        Me.grid_RejectList_Detail.Rows.Count = 2
        Me.grid_RejectList_Detail.Rows.DefaultSize = 20
        Me.grid_RejectList_Detail.Size = New System.Drawing.Size(820, 392)
        Me.grid_RejectList_Detail.StyleInfo = resources.GetString("grid_RejectList_Detail.StyleInfo")
        Me.grid_RejectList_Detail.TabIndex = 22
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Panel3.Controls.Add(Me.date_Reject_End)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.date_Reject_Start)
        Me.Panel3.Controls.Add(Me.rb_Date)
        Me.Panel3.Controls.Add(Me.rb_Area)
        Me.Panel3.Controls.Add(Me.DateTimePicker3)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Controls.Add(Me.btn_InputCheck)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.tb_Max_No)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.tb_Min_No)
        Me.Panel3.Controls.Add(Me.TextBox1)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(839, 142)
        Me.Panel3.TabIndex = 21
        '
        'date_Reject_End
        '
        Me.date_Reject_End.Enabled = False
        Me.date_Reject_End.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.date_Reject_End.Location = New System.Drawing.Point(243, 104)
        Me.date_Reject_End.Margin = New System.Windows.Forms.Padding(0, 3, 3, 3)
        Me.date_Reject_End.Name = "date_Reject_End"
        Me.date_Reject_End.Size = New System.Drawing.Size(101, 21)
        Me.date_Reject_End.TabIndex = 40
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(210, 104)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 21)
        Me.Label2.TabIndex = 39
        Me.Label2.Text = "~"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'date_Reject_Start
        '
        Me.date_Reject_Start.Enabled = False
        Me.date_Reject_Start.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.date_Reject_Start.Location = New System.Drawing.Point(108, 104)
        Me.date_Reject_Start.Margin = New System.Windows.Forms.Padding(0, 3, 3, 3)
        Me.date_Reject_Start.Name = "date_Reject_Start"
        Me.date_Reject_Start.Size = New System.Drawing.Size(101, 21)
        Me.date_Reject_Start.TabIndex = 38
        '
        'rb_Date
        '
        Me.rb_Date.AutoSize = True
        Me.rb_Date.Checked = True
        Me.rb_Date.Location = New System.Drawing.Point(9, 106)
        Me.rb_Date.Name = "rb_Date"
        Me.rb_Date.Size = New System.Drawing.Size(99, 16)
        Me.rb_Date.TabIndex = 37
        Me.rb_Date.TabStop = True
        Me.rb_Date.Text = "회수일자 선택"
        Me.rb_Date.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.rb_Date.UseVisualStyleBackColor = True
        '
        'rb_Area
        '
        Me.rb_Area.AutoSize = True
        Me.rb_Area.Location = New System.Drawing.Point(9, 81)
        Me.rb_Area.Name = "rb_Area"
        Me.rb_Area.Size = New System.Drawing.Size(99, 16)
        Me.rb_Area.TabIndex = 36
        Me.rb_Area.Text = "반납범위 선택"
        Me.rb_Area.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.rb_Area.UseVisualStyleBackColor = True
        '
        'DateTimePicker3
        '
        Me.DateTimePicker3.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker3.Location = New System.Drawing.Point(108, 54)
        Me.DateTimePicker3.Margin = New System.Windows.Forms.Padding(0, 3, 3, 3)
        Me.DateTimePicker3.Name = "DateTimePicker3"
        Me.DateTimePicker3.Size = New System.Drawing.Size(101, 21)
        Me.DateTimePicker3.TabIndex = 35
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(9, 54)
        Me.Label6.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(99, 21)
        Me.Label6.TabIndex = 34
        Me.Label6.Text = "우익반납 일자"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn_InputCheck
        '
        Me.btn_InputCheck.Location = New System.Drawing.Point(475, 93)
        Me.btn_InputCheck.Name = "btn_InputCheck"
        Me.btn_InputCheck.Size = New System.Drawing.Size(145, 43)
        Me.btn_InputCheck.TabIndex = 33
        Me.btn_InputCheck.Text = "자동입력"
        Me.btn_InputCheck.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(280, 77)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(189, 24)
        Me.Label4.TabIndex = 32
        Me.Label4.Text = "범위는 숫자만 입력하여 주십시오." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "예) Y00001 ~ Y00100 = 1 ~ 100"
        '
        'tb_Max_No
        '
        Me.tb_Max_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Max_No.Enabled = False
        Me.tb_Max_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Max_No.Location = New System.Drawing.Point(212, 79)
        Me.tb_Max_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Max_No.Name = "tb_Max_No"
        Me.tb_Max_No.Size = New System.Drawing.Size(65, 21)
        Me.tb_Max_No.TabIndex = 31
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(179, 79)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 21)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "~"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_Min_No
        '
        Me.tb_Min_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Min_No.Enabled = False
        Me.tb_Min_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Min_No.Location = New System.Drawing.Point(111, 79)
        Me.tb_Min_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Min_No.Name = "tb_Min_No"
        Me.tb_Min_No.Size = New System.Drawing.Size(65, 21)
        Me.tb_Min_No.TabIndex = 30
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.BackColor = System.Drawing.SystemColors.Window
        Me.TextBox1.Enabled = False
        Me.TextBox1.Font = New System.Drawing.Font("굴림", 16.0!)
        Me.TextBox1.Location = New System.Drawing.Point(230, 8)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(599, 32)
        Me.TextBox1.TabIndex = 29
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.SteelBlue
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 16.0!)
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(9, 8)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(221, 32)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "반납번호"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me.Label11.Text = "반납 일자"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
        Me.Panel1.Size = New System.Drawing.Size(421, 85)
        Me.Panel1.TabIndex = 6
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.grid_RejectList)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label18)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.White
        Me.SplitContainer1.Panel2.Controls.Add(Me.grid_PartList)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer1.Panel2.Controls.Add(Me.grid_RejectList_Detail)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TextBox2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 748)
        Me.SplitContainer1.SplitterDistance = 421
        Me.SplitContainer1.TabIndex = 10
        '
        'grid_RejectList
        '
        Me.grid_RejectList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_RejectList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_RejectList.Location = New System.Drawing.Point(0, 85)
        Me.grid_RejectList.Name = "grid_RejectList"
        Me.grid_RejectList.Rows.Count = 2
        Me.grid_RejectList.Rows.DefaultSize = 20
        Me.grid_RejectList.Size = New System.Drawing.Size(421, 663)
        Me.grid_RejectList.StyleInfo = resources.GetString("grid_RejectList.StyleInfo")
        Me.grid_RejectList.TabIndex = 5
        '
        'grid_PartList
        '
        Me.grid_PartList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grid_PartList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_PartList.Location = New System.Drawing.Point(9, 604)
        Me.grid_PartList.Name = "grid_PartList"
        Me.grid_PartList.Rows.Count = 2
        Me.grid_PartList.Rows.DefaultSize = 20
        Me.grid_PartList.Size = New System.Drawing.Size(820, 141)
        Me.grid_PartList.StyleInfo = resources.GetString("grid_PartList.StyleInfo")
        Me.grid_PartList.TabIndex = 36
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("굴림", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label12.Location = New System.Drawing.Point(8, 586)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(209, 15)
        Me.Label12.TabIndex = 35
        Me.Label12.Text = "[ Part별 List *전표인쇄영역* ]"
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.SystemColors.Window
        Me.TextBox2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TextBox2.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(0, 722)
        Me.TextBox2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(839, 26)
        Me.TextBox2.TabIndex = 32
        Me.TextBox2.Text = "여기는 누락된 데이터가 표시되는 곳입니다."
        Me.TextBox2.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("굴림", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label10.Location = New System.Drawing.Point(8, 159)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(89, 15)
        Me.Label10.TabIndex = 34
        Me.Label10.Text = "[ 범위 List ]"
        '
        'cms_Menu
        '
        Me.cms_Menu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_RowDelete})
        Me.cms_Menu.Name = "GRID_MENU"
        Me.cms_Menu.Size = New System.Drawing.Size(120, 26)
        '
        'frm_Reject_IC_Ship
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_Reject_IC_Ship"
        Me.Text = "폐기자재 반납"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        CType(Me.grid_RejectList_Detail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grid_RejectList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grid_PartList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cms_Menu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btn_RowDelete As ToolStripMenuItem
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents btn_Search As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents btn_New_Reject As ToolStripButton
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents btn_Save As ToolStripButton
    Friend WithEvents Form_CLose As ToolStripButton
    Friend WithEvents grid_RejectList_Detail As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Label18 As Label
    Friend WithEvents tb_S_YJ_No As TextBox
    Friend WithEvents tb_S_Lot_No As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents DateTimePicker2 As DateTimePicker
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents Label11 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents grid_RejectList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents cms_Menu As ContextMenuStrip
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btn_InputCheck As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents tb_Max_No As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents tb_Min_No As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents DateTimePicker3 As DateTimePicker
    Friend WithEvents Label6 As Label
    Friend WithEvents date_Reject_End As DateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents date_Reject_Start As DateTimePicker
    Friend WithEvents rb_Date As RadioButton
    Friend WithEvents rb_Area As RadioButton
    Friend WithEvents btn_PMIC_Print As ToolStripButton
    Friend WithEvents Label10 As Label
    Friend WithEvents grid_PartList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Label12 As Label
    Friend WithEvents ToolTip1 As ToolTip
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HistoryForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(HistoryForm))
        Me.C1DockingTab1 = New C1.Win.C1Command.C1DockingTab()
        Me.C1DockingTabPage1 = New C1.Win.C1Command.C1DockingTabPage()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Cb_factoryName = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Cb_workSide = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Cb_workLine = New System.Windows.Forms.ComboBox()
        Me.Tb_customerName = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Tb_modelName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Btn_Search = New System.Windows.Forms.Button()
        Me.DTP_endDate = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DTP_startDate = New System.Windows.Forms.DateTimePicker()
        Me.Grid_AllPartsCheck = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.C1DockingTabPage2 = New C1.Win.C1Command.C1DockingTabPage()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.CB_FirstSetting = New System.Windows.Forms.CheckBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Cb_result = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Cb_factoryName2 = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Cb_workSide2 = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Cb_workLine2 = New System.Windows.Forms.ComboBox()
        Me.Tb_customerName2 = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Tb_modelName2 = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.DTP_endDate2 = New System.Windows.Forms.DateTimePicker()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.DTP_startDate2 = New System.Windows.Forms.DateTimePicker()
        Me.Btn_changeSearch = New System.Windows.Forms.Button()
        Me.Grid_PartsChange = New C1.Win.C1FlexGrid.C1FlexGrid()
        CType(Me.C1DockingTab1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.C1DockingTab1.SuspendLayout()
        Me.C1DockingTabPage1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.Grid_AllPartsCheck, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.C1DockingTabPage2.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.Grid_PartsChange, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'C1DockingTab1
        '
        Me.C1DockingTab1.Controls.Add(Me.C1DockingTabPage1)
        Me.C1DockingTab1.Controls.Add(Me.C1DockingTabPage2)
        Me.C1DockingTab1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.C1DockingTab1.HotTrack = True
        Me.C1DockingTab1.Location = New System.Drawing.Point(0, 0)
        Me.C1DockingTab1.Name = "C1DockingTab1"
        Me.C1DockingTab1.SelectedTabBold = True
        Me.C1DockingTab1.Size = New System.Drawing.Size(1264, 773)
        Me.C1DockingTab1.TabIndex = 0
        Me.C1DockingTab1.TabsSpacing = 5
        Me.C1DockingTab1.TabStyle = C1.Win.C1Command.TabStyleEnum.Office2010
        Me.C1DockingTab1.VisualStyle = C1.Win.C1Command.VisualStyle.Office2010Blue
        Me.C1DockingTab1.VisualStyleBase = C1.Win.C1Command.VisualStyle.Office2010Blue
        '
        'C1DockingTabPage1
        '
        Me.C1DockingTabPage1.Controls.Add(Me.SplitContainer1)
        Me.C1DockingTabPage1.Image = Global.YJ_MMS_MMPS.My.Resources.Resources._102_16
        Me.C1DockingTabPage1.Location = New System.Drawing.Point(1, 25)
        Me.C1DockingTabPage1.Name = "C1DockingTabPage1"
        Me.C1DockingTabPage1.Size = New System.Drawing.Size(1262, 747)
        Me.C1DockingTabPage1.TabIndex = 0
        Me.C1DockingTabPage1.Text = "All Parts Check"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Cb_factoryName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Cb_workSide)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Cb_workLine)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Tb_customerName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Tb_modelName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Btn_Search)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DTP_endDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DTP_startDate)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Grid_AllPartsCheck)
        Me.SplitContainer1.Size = New System.Drawing.Size(1262, 747)
        Me.SplitContainer1.SplitterDistance = 64
        Me.SplitContainer1.TabIndex = 0
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(746, 10)
        Me.Label8.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(127, 21)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "공장"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_factoryName
        '
        Me.Cb_factoryName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_factoryName.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_factoryName.FormattingEnabled = True
        Me.Cb_factoryName.Location = New System.Drawing.Point(873, 10)
        Me.Cb_factoryName.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_factoryName.Name = "Cb_factoryName"
        Me.Cb_factoryName.Size = New System.Drawing.Size(141, 21)
        Me.Cb_factoryName.TabIndex = 10
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(1014, 11)
        Me.Label9.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(127, 21)
        Me.Label9.TabIndex = 13
        Me.Label9.Text = "작업면"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_workSide
        '
        Me.Cb_workSide.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_workSide.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_workSide.FormattingEnabled = True
        Me.Cb_workSide.Items.AddRange(New Object() {"Bottom", "Top"})
        Me.Cb_workSide.Location = New System.Drawing.Point(1141, 11)
        Me.Cb_workSide.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_workSide.Name = "Cb_workSide"
        Me.Cb_workSide.Size = New System.Drawing.Size(141, 21)
        Me.Cb_workSide.TabIndex = 14
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label10.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(746, 33)
        Me.Label10.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(127, 21)
        Me.Label10.TabIndex = 11
        Me.Label10.Text = "작업 Line"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_workLine
        '
        Me.Cb_workLine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_workLine.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_workLine.FormattingEnabled = True
        Me.Cb_workLine.Location = New System.Drawing.Point(873, 33)
        Me.Cb_workLine.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_workLine.Name = "Cb_workLine"
        Me.Cb_workLine.Size = New System.Drawing.Size(141, 21)
        Me.Cb_workLine.TabIndex = 12
        '
        'Tb_customerName
        '
        Me.Tb_customerName.BackColor = System.Drawing.SystemColors.Window
        Me.Tb_customerName.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Tb_customerName.Location = New System.Drawing.Point(578, 10)
        Me.Tb_customerName.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_customerName.Name = "Tb_customerName"
        Me.Tb_customerName.Size = New System.Drawing.Size(168, 21)
        Me.Tb_customerName.TabIndex = 6
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(478, 10)
        Me.Label7.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(100, 21)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "고객사"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Tb_modelName
        '
        Me.Tb_modelName.BackColor = System.Drawing.SystemColors.Window
        Me.Tb_modelName.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Tb_modelName.Location = New System.Drawing.Point(578, 33)
        Me.Tb_modelName.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_modelName.Name = "Tb_modelName"
        Me.Tb_modelName.Size = New System.Drawing.Size(168, 21)
        Me.Tb_modelName.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(478, 33)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 21)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "모델명"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Btn_Search
        '
        Me.Btn_Search.Font = New System.Drawing.Font("굴림", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_Search.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.Search
        Me.Btn_Search.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Btn_Search.Location = New System.Drawing.Point(11, 12)
        Me.Btn_Search.Name = "Btn_Search"
        Me.Btn_Search.Size = New System.Drawing.Size(96, 42)
        Me.Btn_Search.TabIndex = 0
        Me.Btn_Search.Text = "조회"
        Me.Btn_Search.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Btn_Search.UseVisualStyleBackColor = True
        '
        'DTP_endDate
        '
        Me.DTP_endDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTP_endDate.Location = New System.Drawing.Point(357, 10)
        Me.DTP_endDate.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DTP_endDate.Name = "DTP_endDate"
        Me.DTP_endDate.Size = New System.Drawing.Size(121, 21)
        Me.DTP_endDate.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(343, 14)
        Me.Label4.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(14, 12)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "~"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(119, 10)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 21)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "조회기간"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DTP_startDate
        '
        Me.DTP_startDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTP_startDate.Location = New System.Drawing.Point(219, 10)
        Me.DTP_startDate.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DTP_startDate.Name = "DTP_startDate"
        Me.DTP_startDate.Size = New System.Drawing.Size(121, 21)
        Me.DTP_startDate.TabIndex = 2
        '
        'Grid_AllPartsCheck
        '
        Me.Grid_AllPartsCheck.AllowFreezing = C1.Win.C1FlexGrid.AllowFreezingEnum.Both
        Me.Grid_AllPartsCheck.AutoResize = True
        Me.Grid_AllPartsCheck.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.Grid_AllPartsCheck.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_AllPartsCheck.Font = New System.Drawing.Font("돋움", 9.0!)
        Me.Grid_AllPartsCheck.KeyActionTab = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross
        Me.Grid_AllPartsCheck.Location = New System.Drawing.Point(0, 0)
        Me.Grid_AllPartsCheck.Name = "Grid_AllPartsCheck"
        Me.Grid_AllPartsCheck.Rows.Count = 2
        Me.Grid_AllPartsCheck.Rows.DefaultSize = 20
        Me.Grid_AllPartsCheck.Size = New System.Drawing.Size(1262, 679)
        Me.Grid_AllPartsCheck.StyleInfo = resources.GetString("Grid_AllPartsCheck.StyleInfo")
        Me.Grid_AllPartsCheck.TabIndex = 0
        Me.Grid_AllPartsCheck.VisualStyle = C1.Win.C1FlexGrid.VisualStyle.Office2010Blue
        '
        'C1DockingTabPage2
        '
        Me.C1DockingTabPage2.Controls.Add(Me.SplitContainer2)
        Me.C1DockingTabPage2.Image = Global.YJ_MMS_MMPS.My.Resources.Resources._121_16
        Me.C1DockingTabPage2.Location = New System.Drawing.Point(1, 25)
        Me.C1DockingTabPage2.Name = "C1DockingTabPage2"
        Me.C1DockingTabPage2.Size = New System.Drawing.Size(1262, 747)
        Me.C1DockingTabPage2.TabIndex = 1
        Me.C1DockingTabPage2.Text = "Parts Change"
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.CB_FirstSetting)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label15)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Cb_result)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Cb_factoryName2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Cb_workSide2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Cb_workLine2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Tb_customerName2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Tb_modelName2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DTP_endDate2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label13)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label14)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DTP_startDate2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Btn_changeSearch)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Grid_PartsChange)
        Me.SplitContainer2.Size = New System.Drawing.Size(1262, 747)
        Me.SplitContainer2.SplitterDistance = 64
        Me.SplitContainer2.TabIndex = 2
        '
        'CB_FirstSetting
        '
        Me.CB_FirstSetting.AutoSize = True
        Me.CB_FirstSetting.Checked = True
        Me.CB_FirstSetting.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CB_FirstSetting.Location = New System.Drawing.Point(119, 38)
        Me.CB_FirstSetting.Name = "CB_FirstSetting"
        Me.CB_FirstSetting.Size = New System.Drawing.Size(119, 16)
        Me.CB_FirstSetting.TabIndex = 31
        Me.CB_FirstSetting.Text = "Setting Data 보기"
        Me.CB_FirstSetting.UseVisualStyleBackColor = True
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label15.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(1014, 33)
        Me.Label15.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(127, 21)
        Me.Label15.TabIndex = 29
        Me.Label15.Text = "결과"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_result
        '
        Me.Cb_result.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_result.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_result.FormattingEnabled = True
        Me.Cb_result.Items.AddRange(New Object() {"ALL", "OK", "NG"})
        Me.Cb_result.Location = New System.Drawing.Point(1141, 33)
        Me.Cb_result.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_result.Name = "Cb_result"
        Me.Cb_result.Size = New System.Drawing.Size(141, 21)
        Me.Cb_result.TabIndex = 30
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(746, 10)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(127, 21)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "공장"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_factoryName2
        '
        Me.Cb_factoryName2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_factoryName2.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_factoryName2.FormattingEnabled = True
        Me.Cb_factoryName2.Location = New System.Drawing.Point(873, 10)
        Me.Cb_factoryName2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_factoryName2.Name = "Cb_factoryName2"
        Me.Cb_factoryName2.Size = New System.Drawing.Size(141, 21)
        Me.Cb_factoryName2.TabIndex = 24
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(1014, 11)
        Me.Label5.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(127, 21)
        Me.Label5.TabIndex = 27
        Me.Label5.Text = "작업면"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_workSide2
        '
        Me.Cb_workSide2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_workSide2.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_workSide2.FormattingEnabled = True
        Me.Cb_workSide2.Items.AddRange(New Object() {"Bottom", "Top"})
        Me.Cb_workSide2.Location = New System.Drawing.Point(1141, 11)
        Me.Cb_workSide2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_workSide2.Name = "Cb_workSide2"
        Me.Cb_workSide2.Size = New System.Drawing.Size(141, 21)
        Me.Cb_workSide2.TabIndex = 28
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(746, 33)
        Me.Label6.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(127, 21)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "작업 Line"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_workLine2
        '
        Me.Cb_workLine2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_workLine2.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_workLine2.FormattingEnabled = True
        Me.Cb_workLine2.Location = New System.Drawing.Point(873, 33)
        Me.Cb_workLine2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_workLine2.Name = "Cb_workLine2"
        Me.Cb_workLine2.Size = New System.Drawing.Size(141, 21)
        Me.Cb_workLine2.TabIndex = 26
        '
        'Tb_customerName2
        '
        Me.Tb_customerName2.BackColor = System.Drawing.SystemColors.Window
        Me.Tb_customerName2.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Tb_customerName2.Location = New System.Drawing.Point(578, 10)
        Me.Tb_customerName2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_customerName2.Name = "Tb_customerName2"
        Me.Tb_customerName2.Size = New System.Drawing.Size(168, 21)
        Me.Tb_customerName2.TabIndex = 20
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label11.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(478, 10)
        Me.Label11.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(100, 21)
        Me.Label11.TabIndex = 19
        Me.Label11.Text = "고객사"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Tb_modelName2
        '
        Me.Tb_modelName2.BackColor = System.Drawing.SystemColors.Window
        Me.Tb_modelName2.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Tb_modelName2.Location = New System.Drawing.Point(578, 33)
        Me.Tb_modelName2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_modelName2.Name = "Tb_modelName2"
        Me.Tb_modelName2.Size = New System.Drawing.Size(168, 21)
        Me.Tb_modelName2.TabIndex = 22
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label12.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label12.ForeColor = System.Drawing.Color.White
        Me.Label12.Location = New System.Drawing.Point(478, 33)
        Me.Label12.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(100, 21)
        Me.Label12.TabIndex = 21
        Me.Label12.Text = "모델명"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DTP_endDate2
        '
        Me.DTP_endDate2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTP_endDate2.Location = New System.Drawing.Point(357, 10)
        Me.DTP_endDate2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DTP_endDate2.Name = "DTP_endDate2"
        Me.DTP_endDate2.Size = New System.Drawing.Size(121, 21)
        Me.DTP_endDate2.TabIndex = 18
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(343, 14)
        Me.Label13.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(14, 12)
        Me.Label13.TabIndex = 17
        Me.Label13.Text = "~"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label14.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(119, 10)
        Me.Label14.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(100, 21)
        Me.Label14.TabIndex = 15
        Me.Label14.Text = "조회기간"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DTP_startDate2
        '
        Me.DTP_startDate2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTP_startDate2.Location = New System.Drawing.Point(219, 10)
        Me.DTP_startDate2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DTP_startDate2.Name = "DTP_startDate2"
        Me.DTP_startDate2.Size = New System.Drawing.Size(121, 21)
        Me.DTP_startDate2.TabIndex = 16
        '
        'Btn_changeSearch
        '
        Me.Btn_changeSearch.Font = New System.Drawing.Font("굴림", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_changeSearch.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.Search
        Me.Btn_changeSearch.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Btn_changeSearch.Location = New System.Drawing.Point(11, 12)
        Me.Btn_changeSearch.Name = "Btn_changeSearch"
        Me.Btn_changeSearch.Size = New System.Drawing.Size(96, 42)
        Me.Btn_changeSearch.TabIndex = 8
        Me.Btn_changeSearch.Text = "조회"
        Me.Btn_changeSearch.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Btn_changeSearch.UseVisualStyleBackColor = True
        '
        'Grid_PartsChange
        '
        Me.Grid_PartsChange.AllowFreezing = C1.Win.C1FlexGrid.AllowFreezingEnum.Both
        Me.Grid_PartsChange.AutoResize = True
        Me.Grid_PartsChange.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.Grid_PartsChange.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_PartsChange.Font = New System.Drawing.Font("돋움", 9.0!)
        Me.Grid_PartsChange.KeyActionTab = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross
        Me.Grid_PartsChange.Location = New System.Drawing.Point(0, 0)
        Me.Grid_PartsChange.Name = "Grid_PartsChange"
        Me.Grid_PartsChange.Rows.Count = 2
        Me.Grid_PartsChange.Rows.DefaultSize = 20
        Me.Grid_PartsChange.Size = New System.Drawing.Size(1262, 679)
        Me.Grid_PartsChange.StyleInfo = resources.GetString("Grid_PartsChange.StyleInfo")
        Me.Grid_PartsChange.TabIndex = 55
        Me.Grid_PartsChange.VisualStyle = C1.Win.C1FlexGrid.VisualStyle.Office2010Blue
        '
        'HistoryForm
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.C1DockingTab1)
        Me.Name = "HistoryForm"
        Me.Text = "이력보기"
        CType(Me.C1DockingTab1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.C1DockingTab1.ResumeLayout(False)
        Me.C1DockingTabPage1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.Grid_AllPartsCheck, System.ComponentModel.ISupportInitialize).EndInit()
        Me.C1DockingTabPage2.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.Grid_PartsChange, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents C1DockingTab1 As C1.Win.C1Command.C1DockingTab
    Friend WithEvents C1DockingTabPage1 As C1.Win.C1Command.C1DockingTabPage
    Friend WithEvents C1DockingTabPage2 As C1.Win.C1Command.C1DockingTabPage
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Btn_Search As Button
    Friend WithEvents DTP_endDate As DateTimePicker
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents DTP_startDate As DateTimePicker
    Friend WithEvents Grid_AllPartsCheck As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents Btn_changeSearch As Button
    Friend WithEvents Grid_PartsChange As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Tb_modelName As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Tb_customerName As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Cb_factoryName As ComboBox
    Friend WithEvents Label9 As Label
    Friend WithEvents Cb_workSide As ComboBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Cb_workLine As ComboBox
    Friend WithEvents Label15 As Label
    Friend WithEvents Cb_result As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Cb_factoryName2 As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Cb_workSide2 As ComboBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Cb_workLine2 As ComboBox
    Friend WithEvents Tb_customerName2 As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents Tb_modelName2 As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents DTP_endDate2 As DateTimePicker
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents DTP_startDate2 As DateTimePicker
    Friend WithEvents CB_FirstSetting As CheckBox
End Class

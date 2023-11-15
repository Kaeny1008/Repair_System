<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BLULabelPrint
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BLULabelPrint))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.GRID_History = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TB_PrintQTY = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TB_StopBits = New System.Windows.Forms.TextBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.TB_Parity = New System.Windows.Forms.TextBox()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.CB_Cable = New System.Windows.Forms.ComboBox()
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
        Me.BTN_LabelPrint = New System.Windows.Forms.Button()
        Me.BTN_PrinterSetting = New System.Windows.Forms.Button()
        Me.TB_HistoryNote = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TB_ModelName = New System.Windows.Forms.TextBox()
        Me.BTN_ModelSearch = New System.Windows.Forms.Button()
        Me.TB_ModelCode = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.BTN_Search = New System.Windows.Forms.ToolStripButton()
        Me.DTP_EndDate = New System.Windows.Forms.DateTimePicker()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.DTP_StartDate = New System.Windows.Forms.DateTimePicker()
        Me.TB_S_ModelName = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.GRID_History, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.White
        Me.SplitContainer1.Panel1.Controls.Add(Me.GRID_History)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.AliceBlue
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BTN_LabelPrint)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BTN_PrinterSetting)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TB_HistoryNote)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TB_ModelName)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BTN_ModelSearch)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TB_ModelCode)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 773)
        Me.SplitContainer1.SplitterDistance = 546
        Me.SplitContainer1.TabIndex = 0
        '
        'GRID_History
        '
        Me.GRID_History.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.GRID_History.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GRID_History.Location = New System.Drawing.Point(0, 83)
        Me.GRID_History.Name = "GRID_History"
        Me.GRID_History.Rows.Count = 2
        Me.GRID_History.Rows.DefaultSize = 20
        Me.GRID_History.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row
        Me.GRID_History.Size = New System.Drawing.Size(546, 690)
        Me.GRID_History.StyleInfo = resources.GetString("GRID_History.StyleInfo")
        Me.GRID_History.TabIndex = 0
        Me.GRID_History.VisualStyle = C1.Win.C1FlexGrid.VisualStyle.Office2010Blue
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.TB_PrintQTY)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.TB_StopBits)
        Me.Panel1.Controls.Add(Me.Label34)
        Me.Panel1.Controls.Add(Me.TB_Parity)
        Me.Panel1.Controls.Add(Me.Label33)
        Me.Panel1.Controls.Add(Me.CB_Cable)
        Me.Panel1.Controls.Add(Me.TB_DataBits)
        Me.Panel1.Controls.Add(Me.Label26)
        Me.Panel1.Controls.Add(Me.Label27)
        Me.Panel1.Controls.Add(Me.Label28)
        Me.Panel1.Controls.Add(Me.TB_TOP_Loc)
        Me.Panel1.Controls.Add(Me.Label29)
        Me.Panel1.Controls.Add(Me.TB_BaudRate)
        Me.Panel1.Controls.Add(Me.TB_LEFT_Loc)
        Me.Panel1.Controls.Add(Me.Label30)
        Me.Panel1.Controls.Add(Me.Label31)
        Me.Panel1.Controls.Add(Me.TB_Port)
        Me.Panel1.Location = New System.Drawing.Point(250, 230)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(345, 271)
        Me.Panel1.TabIndex = 8
        Me.Panel1.Visible = False
        '
        'TB_PrintQTY
        '
        Me.TB_PrintQTY.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_PrintQTY.Location = New System.Drawing.Point(180, 232)
        Me.TB_PrintQTY.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_PrintQTY.Name = "TB_PrintQTY"
        Me.TB_PrintQTY.Size = New System.Drawing.Size(154, 26)
        Me.TB_PrintQTY.TabIndex = 17
        Me.TB_PrintQTY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.SlateGray
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(9, 232)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(171, 26)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "인쇄수량"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_StopBits
        '
        Me.TB_StopBits.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_StopBits.Location = New System.Drawing.Point(180, 204)
        Me.TB_StopBits.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_StopBits.Name = "TB_StopBits"
        Me.TB_StopBits.Size = New System.Drawing.Size(154, 26)
        Me.TB_StopBits.TabIndex = 15
        Me.TB_StopBits.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label34
        '
        Me.Label34.BackColor = System.Drawing.Color.SlateGray
        Me.Label34.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label34.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label34.ForeColor = System.Drawing.Color.White
        Me.Label34.Location = New System.Drawing.Point(9, 204)
        Me.Label34.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(171, 26)
        Me.Label34.TabIndex = 14
        Me.Label34.Text = "StopBits"
        Me.Label34.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Parity
        '
        Me.TB_Parity.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_Parity.Location = New System.Drawing.Point(180, 176)
        Me.TB_Parity.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Parity.Name = "TB_Parity"
        Me.TB_Parity.Size = New System.Drawing.Size(154, 26)
        Me.TB_Parity.TabIndex = 13
        Me.TB_Parity.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label33
        '
        Me.Label33.BackColor = System.Drawing.Color.SlateGray
        Me.Label33.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label33.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label33.ForeColor = System.Drawing.Color.White
        Me.Label33.Location = New System.Drawing.Point(9, 176)
        Me.Label33.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(171, 26)
        Me.Label33.TabIndex = 12
        Me.Label33.Text = "Parity"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CB_Cable
        '
        Me.CB_Cable.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_Cable.Font = New System.Drawing.Font("굴림", 13.75!, System.Drawing.FontStyle.Bold)
        Me.CB_Cable.FormattingEnabled = True
        Me.CB_Cable.Items.AddRange(New Object() {"COM", "LPT"})
        Me.CB_Cable.Location = New System.Drawing.Point(180, 8)
        Me.CB_Cable.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.CB_Cable.Name = "CB_Cable"
        Me.CB_Cable.Size = New System.Drawing.Size(154, 26)
        Me.CB_Cable.TabIndex = 1
        '
        'TB_DataBits
        '
        Me.TB_DataBits.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_DataBits.Location = New System.Drawing.Point(180, 148)
        Me.TB_DataBits.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_DataBits.Name = "TB_DataBits"
        Me.TB_DataBits.Size = New System.Drawing.Size(154, 26)
        Me.TB_DataBits.TabIndex = 11
        Me.TB_DataBits.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.Color.SlateGray
        Me.Label26.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label26.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.White
        Me.Label26.Location = New System.Drawing.Point(9, 8)
        Me.Label26.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(171, 26)
        Me.Label26.TabIndex = 0
        Me.Label26.Text = "COM / LPT"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label27
        '
        Me.Label27.BackColor = System.Drawing.Color.SlateGray
        Me.Label27.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label27.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label27.ForeColor = System.Drawing.Color.White
        Me.Label27.Location = New System.Drawing.Point(9, 148)
        Me.Label27.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(171, 26)
        Me.Label27.TabIndex = 10
        Me.Label27.Text = "DataBits"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.Color.SlateGray
        Me.Label28.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label28.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label28.ForeColor = System.Drawing.Color.White
        Me.Label28.Location = New System.Drawing.Point(9, 92)
        Me.Label28.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(171, 26)
        Me.Label28.TabIndex = 6
        Me.Label28.Text = "인쇄위치(세로)"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_TOP_Loc
        '
        Me.TB_TOP_Loc.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_TOP_Loc.Location = New System.Drawing.Point(180, 92)
        Me.TB_TOP_Loc.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_TOP_Loc.Name = "TB_TOP_Loc"
        Me.TB_TOP_Loc.Size = New System.Drawing.Size(154, 26)
        Me.TB_TOP_Loc.TabIndex = 7
        Me.TB_TOP_Loc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label29
        '
        Me.Label29.BackColor = System.Drawing.Color.SlateGray
        Me.Label29.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label29.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label29.ForeColor = System.Drawing.Color.White
        Me.Label29.Location = New System.Drawing.Point(9, 120)
        Me.Label29.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(171, 26)
        Me.Label29.TabIndex = 8
        Me.Label29.Text = "BaudRate"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_BaudRate
        '
        Me.TB_BaudRate.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_BaudRate.Location = New System.Drawing.Point(180, 120)
        Me.TB_BaudRate.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_BaudRate.Name = "TB_BaudRate"
        Me.TB_BaudRate.Size = New System.Drawing.Size(154, 26)
        Me.TB_BaudRate.TabIndex = 9
        Me.TB_BaudRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TB_LEFT_Loc
        '
        Me.TB_LEFT_Loc.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_LEFT_Loc.Location = New System.Drawing.Point(180, 64)
        Me.TB_LEFT_Loc.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_LEFT_Loc.Name = "TB_LEFT_Loc"
        Me.TB_LEFT_Loc.Size = New System.Drawing.Size(154, 26)
        Me.TB_LEFT_Loc.TabIndex = 5
        Me.TB_LEFT_Loc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label30
        '
        Me.Label30.BackColor = System.Drawing.Color.SlateGray
        Me.Label30.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label30.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label30.ForeColor = System.Drawing.Color.White
        Me.Label30.Location = New System.Drawing.Point(9, 36)
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
        Me.Label31.Location = New System.Drawing.Point(9, 64)
        Me.Label31.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(171, 26)
        Me.Label31.TabIndex = 4
        Me.Label31.Text = "인쇄위치(가로)"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_Port
        '
        Me.TB_Port.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TB_Port.Location = New System.Drawing.Point(180, 36)
        Me.TB_Port.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Port.Name = "TB_Port"
        Me.TB_Port.Size = New System.Drawing.Size(154, 26)
        Me.TB_Port.TabIndex = 3
        Me.TB_Port.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'BTN_LabelPrint
        '
        Me.BTN_LabelPrint.Font = New System.Drawing.Font("굴림", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_LabelPrint.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.barcode
        Me.BTN_LabelPrint.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BTN_LabelPrint.Location = New System.Drawing.Point(457, 167)
        Me.BTN_LabelPrint.Name = "BTN_LabelPrint"
        Me.BTN_LabelPrint.Size = New System.Drawing.Size(147, 57)
        Me.BTN_LabelPrint.TabIndex = 7
        Me.BTN_LabelPrint.Text = "라벨발행"
        Me.BTN_LabelPrint.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BTN_LabelPrint.UseVisualStyleBackColor = True
        '
        'BTN_PrinterSetting
        '
        Me.BTN_PrinterSetting.Font = New System.Drawing.Font("굴림", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_PrinterSetting.Location = New System.Drawing.Point(250, 167)
        Me.BTN_PrinterSetting.Name = "BTN_PrinterSetting"
        Me.BTN_PrinterSetting.Size = New System.Drawing.Size(200, 57)
        Me.BTN_PrinterSetting.TabIndex = 6
        Me.BTN_PrinterSetting.Text = "프린터 설정 보기"
        Me.BTN_PrinterSetting.UseVisualStyleBackColor = True
        '
        'TB_HistoryNote
        '
        Me.TB_HistoryNote.Font = New System.Drawing.Font("굴림", 12.0!)
        Me.TB_HistoryNote.Location = New System.Drawing.Point(250, 56)
        Me.TB_HistoryNote.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_HistoryNote.Multiline = True
        Me.TB_HistoryNote.Name = "TB_HistoryNote"
        Me.TB_HistoryNote.Size = New System.Drawing.Size(354, 107)
        Me.TB_HistoryNote.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 12.0!)
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(17, 56)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(233, 107)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "비고"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TB_ModelName
        '
        Me.TB_ModelName.Font = New System.Drawing.Font("굴림", 12.0!)
        Me.TB_ModelName.Location = New System.Drawing.Point(350, 25)
        Me.TB_ModelName.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_ModelName.Name = "TB_ModelName"
        Me.TB_ModelName.Size = New System.Drawing.Size(254, 26)
        Me.TB_ModelName.TabIndex = 2
        '
        'BTN_ModelSearch
        '
        Me.BTN_ModelSearch.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.Search_12
        Me.BTN_ModelSearch.Location = New System.Drawing.Point(604, 24)
        Me.BTN_ModelSearch.Margin = New System.Windows.Forms.Padding(0)
        Me.BTN_ModelSearch.Name = "BTN_ModelSearch"
        Me.BTN_ModelSearch.Size = New System.Drawing.Size(23, 28)
        Me.BTN_ModelSearch.TabIndex = 3
        Me.BTN_ModelSearch.UseVisualStyleBackColor = True
        '
        'TB_ModelCode
        '
        Me.TB_ModelCode.Enabled = False
        Me.TB_ModelCode.Font = New System.Drawing.Font("굴림", 12.0!)
        Me.TB_ModelCode.Location = New System.Drawing.Point(250, 25)
        Me.TB_ModelCode.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_ModelCode.Name = "TB_ModelCode"
        Me.TB_ModelCode.Size = New System.Drawing.Size(100, 26)
        Me.TB_ModelCode.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 12.0!)
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(17, 25)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(233, 26)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Model CODE / Model Name"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(348, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(269, 12)
        Me.Label4.TabIndex = 127
        Me.Label4.Text = "※ 영문 대소문자를 구분하여 검색하여 주십시오."
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Panel2.Controls.Add(Me.TB_S_ModelName)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.DTP_EndDate)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.DTP_StartDate)
        Me.Panel2.Controls.Add(Me.ToolStrip1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(546, 83)
        Me.Panel2.TabIndex = 1
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_Search})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(546, 25)
        Me.ToolStrip1.TabIndex = 0
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'BTN_Search
        '
        Me.BTN_Search.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.Search
        Me.BTN_Search.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_Search.Name = "BTN_Search"
        Me.BTN_Search.Size = New System.Drawing.Size(51, 22)
        Me.BTN_Search.Text = "검색"
        '
        'DTP_EndDate
        '
        Me.DTP_EndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTP_EndDate.Location = New System.Drawing.Point(247, 30)
        Me.DTP_EndDate.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DTP_EndDate.Name = "DTP_EndDate"
        Me.DTP_EndDate.Size = New System.Drawing.Size(121, 21)
        Me.DTP_EndDate.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(233, 34)
        Me.Label5.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(14, 12)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "~"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(9, 30)
        Me.Label6.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 21)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "발행일자"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DTP_StartDate
        '
        Me.DTP_StartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTP_StartDate.Location = New System.Drawing.Point(109, 30)
        Me.DTP_StartDate.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DTP_StartDate.Name = "DTP_StartDate"
        Me.DTP_StartDate.Size = New System.Drawing.Size(121, 21)
        Me.DTP_StartDate.TabIndex = 6
        '
        'TB_S_ModelName
        '
        Me.TB_S_ModelName.BackColor = System.Drawing.SystemColors.Window
        Me.TB_S_ModelName.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_S_ModelName.Location = New System.Drawing.Point(109, 53)
        Me.TB_S_ModelName.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_S_ModelName.Name = "TB_S_ModelName"
        Me.TB_S_ModelName.Size = New System.Drawing.Size(259, 21)
        Me.TB_S_ModelName.TabIndex = 10
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(9, 53)
        Me.Label7.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(100, 21)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "모델명"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BLULabelPrint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "BLULabelPrint"
        Me.Text = "Label 발행"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.GRID_History, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Panel1 As Panel
    Friend WithEvents TB_StopBits As TextBox
    Friend WithEvents Label34 As Label
    Friend WithEvents TB_Parity As TextBox
    Friend WithEvents Label33 As Label
    Friend WithEvents CB_Cable As ComboBox
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
    Friend WithEvents TB_PrintQTY As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents GRID_History As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents TB_ModelCode As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents BTN_ModelSearch As Button
    Friend WithEvents TB_ModelName As TextBox
    Friend WithEvents BTN_PrinterSetting As Button
    Friend WithEvents TB_HistoryNote As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents BTN_LabelPrint As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents BTN_Search As ToolStripButton
    Friend WithEvents DTP_EndDate As DateTimePicker
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents DTP_StartDate As DateTimePicker
    Friend WithEvents TB_S_ModelName As TextBox
    Friend WithEvents Label7 As Label
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DeviceData
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DeviceData))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Cb_FactoryName = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Cb_workSide = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Cb_workLine = New System.Windows.Forms.ComboBox()
        Me.Btn_newModel = New System.Windows.Forms.Button()
        Me.Tb_modelCode = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Cb_modelName = New System.Windows.Forms.ComboBox()
        Me.Btn_newCustomer = New System.Windows.Forms.Button()
        Me.Tb_customerCode = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Cb_customerName = New System.Windows.Forms.ComboBox()
        Me.Grid_DeviceData = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.BTN_Search = New System.Windows.Forms.ToolStripButton()
        Me.TB_MakerCheck = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_MX_Feeder = New System.Windows.Forms.ToolStripButton()
        Me.BTN_SM_Feeder = New System.Windows.Forms.ToolStripButton()
        Me.BTN_SM_Feeder2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.CMS_gridMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.BTN_newLine = New System.Windows.Forms.ToolStripMenuItem()
        Me.BTN_deleteLine = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_Save = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_DataPrint = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label5 = New System.Windows.Forms.Label()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.Grid_DeviceData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.CMS_gridMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Cb_FactoryName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Cb_workSide)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Cb_workLine)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Btn_newModel)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Tb_modelCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Cb_modelName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Btn_newCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Tb_customerCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Cb_customerName)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Grid_DeviceData)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 748)
        Me.SplitContainer1.SplitterDistance = 62
        Me.SplitContainer1.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(588, 10)
        Me.Label4.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(127, 21)
        Me.Label4.TabIndex = 32
        Me.Label4.Text = "공장"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_FactoryName
        '
        Me.Cb_FactoryName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_FactoryName.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_FactoryName.FormattingEnabled = True
        Me.Cb_FactoryName.Location = New System.Drawing.Point(715, 10)
        Me.Cb_FactoryName.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_FactoryName.Name = "Cb_FactoryName"
        Me.Cb_FactoryName.Size = New System.Drawing.Size(141, 21)
        Me.Cb_FactoryName.TabIndex = 33
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(856, 11)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(127, 21)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "작업면"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_workSide
        '
        Me.Cb_workSide.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_workSide.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_workSide.FormattingEnabled = True
        Me.Cb_workSide.Items.AddRange(New Object() {"Bottom", "Top"})
        Me.Cb_workSide.Location = New System.Drawing.Point(983, 11)
        Me.Cb_workSide.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_workSide.Name = "Cb_workSide"
        Me.Cb_workSide.Size = New System.Drawing.Size(141, 21)
        Me.Cb_workSide.TabIndex = 31
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(588, 33)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(127, 21)
        Me.Label2.TabIndex = 28
        Me.Label2.Text = "작업 Line"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_workLine
        '
        Me.Cb_workLine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_workLine.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_workLine.FormattingEnabled = True
        Me.Cb_workLine.Location = New System.Drawing.Point(715, 33)
        Me.Cb_workLine.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_workLine.Name = "Cb_workLine"
        Me.Cb_workLine.Size = New System.Drawing.Size(141, 21)
        Me.Cb_workLine.TabIndex = 29
        '
        'Btn_newModel
        '
        Me.Btn_newModel.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.plus_blue
        Me.Btn_newModel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Btn_newModel.Location = New System.Drawing.Point(528, 32)
        Me.Btn_newModel.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Btn_newModel.Name = "Btn_newModel"
        Me.Btn_newModel.Size = New System.Drawing.Size(60, 23)
        Me.Btn_newModel.TabIndex = 27
        Me.Btn_newModel.Text = "신규"
        Me.Btn_newModel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Btn_newModel.UseVisualStyleBackColor = True
        '
        'Tb_modelCode
        '
        Me.Tb_modelCode.BackColor = System.Drawing.SystemColors.Window
        Me.Tb_modelCode.Enabled = False
        Me.Tb_modelCode.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Tb_modelCode.Location = New System.Drawing.Point(136, 33)
        Me.Tb_modelCode.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_modelCode.Name = "Tb_modelCode"
        Me.Tb_modelCode.Size = New System.Drawing.Size(75, 21)
        Me.Tb_modelCode.TabIndex = 26
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(9, 33)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(127, 21)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "모델명"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_modelName
        '
        Me.Cb_modelName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_modelName.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_modelName.FormattingEnabled = True
        Me.Cb_modelName.Location = New System.Drawing.Point(211, 33)
        Me.Cb_modelName.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_modelName.Name = "Cb_modelName"
        Me.Cb_modelName.Size = New System.Drawing.Size(317, 21)
        Me.Cb_modelName.TabIndex = 25
        '
        'Btn_newCustomer
        '
        Me.Btn_newCustomer.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.plus_blue
        Me.Btn_newCustomer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Btn_newCustomer.Location = New System.Drawing.Point(528, 9)
        Me.Btn_newCustomer.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Btn_newCustomer.Name = "Btn_newCustomer"
        Me.Btn_newCustomer.Size = New System.Drawing.Size(60, 23)
        Me.Btn_newCustomer.TabIndex = 23
        Me.Btn_newCustomer.Text = "신규"
        Me.Btn_newCustomer.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Btn_newCustomer.UseVisualStyleBackColor = True
        '
        'Tb_customerCode
        '
        Me.Tb_customerCode.BackColor = System.Drawing.SystemColors.Window
        Me.Tb_customerCode.Enabled = False
        Me.Tb_customerCode.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Tb_customerCode.Location = New System.Drawing.Point(136, 10)
        Me.Tb_customerCode.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_customerCode.Name = "Tb_customerCode"
        Me.Tb_customerCode.Size = New System.Drawing.Size(75, 21)
        Me.Tb_customerCode.TabIndex = 22
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(9, 10)
        Me.Label7.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(127, 21)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "고객사"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_customerName
        '
        Me.Cb_customerName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_customerName.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_customerName.FormattingEnabled = True
        Me.Cb_customerName.Location = New System.Drawing.Point(211, 10)
        Me.Cb_customerName.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_customerName.Name = "Cb_customerName"
        Me.Cb_customerName.Size = New System.Drawing.Size(317, 21)
        Me.Cb_customerName.TabIndex = 12
        '
        'Grid_DeviceData
        '
        Me.Grid_DeviceData.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.Grid_DeviceData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_DeviceData.Location = New System.Drawing.Point(0, 0)
        Me.Grid_DeviceData.Name = "Grid_DeviceData"
        Me.Grid_DeviceData.Rows.Count = 2
        Me.Grid_DeviceData.Rows.DefaultSize = 20
        Me.Grid_DeviceData.Size = New System.Drawing.Size(1264, 682)
        Me.Grid_DeviceData.StyleInfo = resources.GetString("Grid_DeviceData.StyleInfo")
        Me.Grid_DeviceData.TabIndex = 4
        Me.Grid_DeviceData.VisualStyle = C1.Win.C1FlexGrid.VisualStyle.Office2010Blue
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_Search, Me.TB_MakerCheck, Me.ToolStripSeparator3, Me.BTN_MX_Feeder, Me.BTN_SM_Feeder, Me.BTN_SM_Feeder2, Me.ToolStripButton1})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1264, 25)
        Me.ToolStrip1.TabIndex = 1
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
        'TB_MakerCheck
        '
        Me.TB_MakerCheck.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.exchange
        Me.TB_MakerCheck.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TB_MakerCheck.Name = "TB_MakerCheck"
        Me.TB_MakerCheck.Size = New System.Drawing.Size(119, 22)
        Me.TB_MakerCheck.Text = "자재 제조사 확인"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
        '
        'BTN_MX_Feeder
        '
        Me.BTN_MX_Feeder.Image = CType(resources.GetObject("BTN_MX_Feeder.Image"), System.Drawing.Image)
        Me.BTN_MX_Feeder.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_MX_Feeder.Name = "BTN_MX_Feeder"
        Me.BTN_MX_Feeder.Size = New System.Drawing.Size(157, 22)
        Me.BTN_MX_Feeder.Text = "MIRAE Feeder List(*.prg)"
        '
        'BTN_SM_Feeder
        '
        Me.BTN_SM_Feeder.Image = CType(resources.GetObject("BTN_SM_Feeder.Image"), System.Drawing.Image)
        Me.BTN_SM_Feeder.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_SM_Feeder.Name = "BTN_SM_Feeder"
        Me.BTN_SM_Feeder.Size = New System.Drawing.Size(244, 22)
        Me.BTN_SM_Feeder.Text = "SAMSUNG(HANWHA) Feeder List(*.xml)"
        '
        'BTN_SM_Feeder2
        '
        Me.BTN_SM_Feeder2.Image = CType(resources.GetObject("BTN_SM_Feeder2.Image"), System.Drawing.Image)
        Me.BTN_SM_Feeder2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_SM_Feeder2.Name = "BTN_SM_Feeder2"
        Me.BTN_SM_Feeder2.Size = New System.Drawing.Size(241, 22)
        Me.BTN_SM_Feeder2.Text = "SAMSUNG(HANWHA) Feeder List(*.csv)"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(127, 22)
        Me.ToolStripButton1.Text = "SAMSUNG CAD Data"
        Me.ToolStripButton1.Visible = False
        '
        'CMS_gridMenu
        '
        Me.CMS_gridMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_newLine, Me.BTN_deleteLine, Me.ToolStripSeparator1, Me.BTN_Save, Me.ToolStripSeparator2, Me.BTN_DataPrint})
        Me.CMS_gridMenu.Name = "CMS_gridMenu"
        Me.CMS_gridMenu.Size = New System.Drawing.Size(164, 104)
        '
        'BTN_newLine
        '
        Me.BTN_newLine.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.plus_blue
        Me.BTN_newLine.Name = "BTN_newLine"
        Me.BTN_newLine.Size = New System.Drawing.Size(163, 22)
        Me.BTN_newLine.Text = "신규등록"
        '
        'BTN_deleteLine
        '
        Me.BTN_deleteLine.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.minus_blue
        Me.BTN_deleteLine.Name = "BTN_deleteLine"
        Me.BTN_deleteLine.Size = New System.Drawing.Size(163, 22)
        Me.BTN_deleteLine.Text = "삭제"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(160, 6)
        '
        'BTN_Save
        '
        Me.BTN_Save.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.save_5
        Me.BTN_Save.Name = "BTN_Save"
        Me.BTN_Save.Size = New System.Drawing.Size(163, 22)
        Me.BTN_Save.Text = "저장"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(160, 6)
        '
        'BTN_DataPrint
        '
        Me.BTN_DataPrint.Name = "BTN_DataPrint"
        Me.BTN_DataPrint.Size = New System.Drawing.Size(163, 22)
        Me.BTN_DataPrint.Text = "DeviceData 출력"
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(1033, 43)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(228, 12)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "(F1) 입력시 신규등록 항목이 실행됩니다."
        '
        'DeviceData
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "DeviceData"
        Me.Text = "DeviceData 작성"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.Grid_DeviceData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.CMS_gridMenu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Grid_DeviceData As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Label7 As Label
    Friend WithEvents Cb_customerName As ComboBox
    Friend WithEvents Tb_customerCode As TextBox
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents BTN_Search As ToolStripButton
    Friend WithEvents Btn_newCustomer As Button
    Friend WithEvents Btn_newModel As Button
    Friend WithEvents Tb_modelCode As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Cb_modelName As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Cb_workLine As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Cb_workSide As ComboBox
    Friend WithEvents CMS_gridMenu As ContextMenuStrip
    Friend WithEvents BTN_newLine As ToolStripMenuItem
    Friend WithEvents BTN_deleteLine As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents BTN_Save As ToolStripMenuItem
    Friend WithEvents BTN_SM_Feeder As ToolStripButton
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents BTN_DataPrint As ToolStripMenuItem
    Friend WithEvents Label4 As Label
    Friend WithEvents Cb_FactoryName As ComboBox
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents BTN_MX_Feeder As ToolStripButton
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents BTN_SM_Feeder2 As ToolStripButton
    Friend WithEvents TB_MakerCheck As ToolStripButton
    Friend WithEvents Label5 As Label
End Class

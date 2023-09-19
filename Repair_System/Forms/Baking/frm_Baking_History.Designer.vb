<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_Baking_History
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Baking_History))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.End_Date = New System.Windows.Forms.DateTimePicker()
        Me.Start_Date = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.BTN_Search = New System.Windows.Forms.ToolStripButton()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.grid_DailyList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.grid_LotList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.tb_Lot_No = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.grid_DailyList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grid_LotList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.tb_Lot_No)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.End_Date)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Start_Date)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ToolStrip1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 773)
        Me.SplitContainer1.SplitterDistance = 70
        Me.SplitContainer1.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(190, 36)
        Me.Label5.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 21)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "~"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'End_Date
        '
        Me.End_Date.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.End_Date.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.End_Date.Location = New System.Drawing.Point(220, 36)
        Me.End_Date.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.End_Date.Name = "End_Date"
        Me.End_Date.Size = New System.Drawing.Size(103, 21)
        Me.End_Date.TabIndex = 7
        '
        'Start_Date
        '
        Me.Start_Date.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Start_Date.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.Start_Date.Location = New System.Drawing.Point(87, 36)
        Me.Start_Date.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Start_Date.Name = "Start_Date"
        Me.Start_Date.Size = New System.Drawing.Size(103, 21)
        Me.Start_Date.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(9, 36)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 21)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "작업 일자"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_Search, Me.Form_CLose})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1264, 25)
        Me.ToolStrip1.TabIndex = 1
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
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.grid_DailyList)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.grid_LotList)
        Me.SplitContainer2.Size = New System.Drawing.Size(1264, 699)
        Me.SplitContainer2.SplitterDistance = 312
        Me.SplitContainer2.TabIndex = 2
        '
        'grid_DailyList
        '
        Me.grid_DailyList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_DailyList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_DailyList.Location = New System.Drawing.Point(0, 0)
        Me.grid_DailyList.Name = "grid_DailyList"
        Me.grid_DailyList.Rows.Count = 2
        Me.grid_DailyList.Rows.DefaultSize = 20
        Me.grid_DailyList.Size = New System.Drawing.Size(312, 699)
        Me.grid_DailyList.StyleInfo = resources.GetString("grid_DailyList.StyleInfo")
        Me.grid_DailyList.TabIndex = 3
        '
        'grid_LotList
        '
        Me.grid_LotList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_LotList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_LotList.Location = New System.Drawing.Point(0, 0)
        Me.grid_LotList.Name = "grid_LotList"
        Me.grid_LotList.Rows.Count = 2
        Me.grid_LotList.Rows.DefaultSize = 20
        Me.grid_LotList.Size = New System.Drawing.Size(948, 699)
        Me.grid_LotList.StyleInfo = resources.GetString("grid_LotList.StyleInfo")
        Me.grid_LotList.TabIndex = 2
        '
        'tb_Lot_No
        '
        Me.tb_Lot_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Lot_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Lot_No.Location = New System.Drawing.Point(401, 36)
        Me.tb_Lot_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Lot_No.Name = "tb_Lot_No"
        Me.tb_Lot_No.Size = New System.Drawing.Size(242, 21)
        Me.tb_Lot_No.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(323, 36)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 21)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Lot No."
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frm_Baking_History
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "frm_Baking_History"
        Me.Text = "Baking 작업현황"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.grid_DailyList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grid_LotList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Label5 As Label
    Friend WithEvents End_Date As DateTimePicker
    Friend WithEvents Start_Date As DateTimePicker
    Friend WithEvents Label1 As Label
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents BTN_Search As ToolStripButton
    Friend WithEvents Form_CLose As ToolStripButton
    Friend WithEvents grid_LotList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents grid_DailyList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents tb_Lot_No As TextBox
    Friend WithEvents Label2 As Label
End Class

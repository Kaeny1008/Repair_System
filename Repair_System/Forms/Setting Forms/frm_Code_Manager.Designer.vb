<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_code_manager
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_code_manager))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.grid_DivisionList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.grid_DetailList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.cms_Grid = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.btn_Add = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_Delete = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Save = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.BTN_Search = New System.Windows.Forms.ToolStripButton()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grid_DivisionList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grid_DetailList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cms_Grid.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.grid_DivisionList)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.grid_DetailList)
        Me.SplitContainer1.Size = New System.Drawing.Size(1138, 687)
        Me.SplitContainer1.SplitterDistance = 412
        Me.SplitContainer1.TabIndex = 1
        '
        'grid_DivisionList
        '
        Me.grid_DivisionList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_DivisionList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_DivisionList.Location = New System.Drawing.Point(0, 48)
        Me.grid_DivisionList.Name = "grid_DivisionList"
        Me.grid_DivisionList.Rows.Count = 2
        Me.grid_DivisionList.Rows.DefaultSize = 20
        Me.grid_DivisionList.Size = New System.Drawing.Size(412, 639)
        Me.grid_DivisionList.StyleInfo = resources.GetString("grid_DivisionList.StyleInfo")
        Me.grid_DivisionList.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label1.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold)
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(412, 48)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "분류"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'grid_DetailList
        '
        Me.grid_DetailList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_DetailList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_DetailList.Location = New System.Drawing.Point(0, 0)
        Me.grid_DetailList.Name = "grid_DetailList"
        Me.grid_DetailList.Rows.Count = 2
        Me.grid_DetailList.Rows.DefaultSize = 20
        Me.grid_DetailList.Size = New System.Drawing.Size(722, 687)
        Me.grid_DetailList.StyleInfo = resources.GetString("grid_DetailList.StyleInfo")
        Me.grid_DetailList.TabIndex = 4
        '
        'cms_Grid
        '
        Me.cms_Grid.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Add, Me.btn_Delete, Me.ToolStripSeparator1, Me.btn_Save})
        Me.cms_Grid.Name = "GRID_MENU"
        Me.cms_Grid.Size = New System.Drawing.Size(99, 76)
        '
        'btn_Add
        '
        Me.btn_Add.Name = "btn_Add"
        Me.btn_Add.Size = New System.Drawing.Size(98, 22)
        Me.btn_Add.Text = "추가"
        '
        'btn_Delete
        '
        Me.btn_Delete.Name = "btn_Delete"
        Me.btn_Delete.Size = New System.Drawing.Size(98, 22)
        Me.btn_Delete.Text = "삭제"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(95, 6)
        '
        'btn_Save
        '
        Me.btn_Save.Name = "btn_Save"
        Me.btn_Save.Size = New System.Drawing.Size(98, 22)
        Me.btn_Save.Text = "저장"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_Search, Me.Form_CLose})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1138, 25)
        Me.ToolStrip1.TabIndex = 6
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
        'frm_code_manager
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1138, 712)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_code_manager"
        Me.Text = "frm_CodeManager"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grid_DivisionList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grid_DetailList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cms_Grid.ResumeLayout(False)
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Label1 As Label
    Friend WithEvents cms_Grid As ContextMenuStrip
    Friend WithEvents btn_Add As ToolStripMenuItem
    Friend WithEvents btn_Delete As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents btn_Save As ToolStripMenuItem
    Friend WithEvents grid_DivisionList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents grid_DetailList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents BTN_Search As ToolStripButton
    Friend WithEvents Form_CLose As ToolStripButton
End Class

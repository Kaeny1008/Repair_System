<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_PG_Upload
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_PG_Upload))
        Me.grid_FileList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.cms_Grid = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.BTN_ADD = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_DELETE = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_SAVE = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_REFRESH = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        Me.BTN_Search = New System.Windows.Forms.ToolStripButton()
        CType(Me.grid_FileList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cms_Grid.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'grid_FileList
        '
        Me.grid_FileList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_FileList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_FileList.Location = New System.Drawing.Point(0, 73)
        Me.grid_FileList.Name = "grid_FileList"
        Me.grid_FileList.Rows.Count = 2
        Me.grid_FileList.Rows.DefaultSize = 20
        Me.grid_FileList.Size = New System.Drawing.Size(999, 559)
        Me.grid_FileList.StyleInfo = resources.GetString("grid_FileList.StyleInfo")
        Me.grid_FileList.TabIndex = 3
        '
        'cms_Grid
        '
        Me.cms_Grid.Font = New System.Drawing.Font("Arial", 9.0!)
        Me.cms_Grid.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_ADD, Me.ToolStripSeparator1, Me.BTN_DELETE, Me.ToolStripSeparator2, Me.BTN_SAVE, Me.ToolStripSeparator3, Me.BTN_REFRESH})
        Me.cms_Grid.Name = "GRID_MENU"
        Me.cms_Grid.Size = New System.Drawing.Size(126, 110)
        '
        'BTN_ADD
        '
        Me.BTN_ADD.Name = "BTN_ADD"
        Me.BTN_ADD.Size = New System.Drawing.Size(125, 22)
        Me.BTN_ADD.Text = "모듈 추가"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(122, 6)
        '
        'BTN_DELETE
        '
        Me.BTN_DELETE.Name = "BTN_DELETE"
        Me.BTN_DELETE.Size = New System.Drawing.Size(125, 22)
        Me.BTN_DELETE.Text = "모듈삭제"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(122, 6)
        '
        'BTN_SAVE
        '
        Me.BTN_SAVE.Name = "BTN_SAVE"
        Me.BTN_SAVE.Size = New System.Drawing.Size(125, 22)
        Me.BTN_SAVE.Text = "저장"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(122, 6)
        '
        'BTN_REFRESH
        '
        Me.BTN_REFRESH.Name = "BTN_REFRESH"
        Me.BTN_REFRESH.Size = New System.Drawing.Size(125, 22)
        Me.BTN_REFRESH.Text = "새로고침"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label1.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold)
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(0, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(999, 48)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "프로그램 파일 관리"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_Search, Me.Form_CLose})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(999, 25)
        Me.ToolStrip1.TabIndex = 5
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
        'BTN_Search
        '
        Me.BTN_Search.Image = Global.Repair_System.My.Resources.Resources.Search_121
        Me.BTN_Search.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_Search.Name = "BTN_Search"
        Me.BTN_Search.Size = New System.Drawing.Size(51, 22)
        Me.BTN_Search.Text = "조회"
        '
        'frm_PG_Upload
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(999, 632)
        Me.Controls.Add(Me.grid_FileList)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_PG_Upload"
        Me.Text = "프로그램 파일 Upload"
        CType(Me.grid_FileList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cms_Grid.ResumeLayout(False)
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grid_FileList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents BTN_ADD As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents BTN_DELETE As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents BTN_SAVE As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents BTN_REFRESH As ToolStripMenuItem
    Friend WithEvents cms_Grid As ContextMenuStrip
    Friend WithEvents Label1 As Label
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents Form_CLose As ToolStripButton
    Friend WithEvents BTN_Search As ToolStripButton
End Class

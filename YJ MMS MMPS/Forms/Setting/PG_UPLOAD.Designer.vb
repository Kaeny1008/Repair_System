<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PG_UPLOAD
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PG_UPLOAD))
        Me.FILE_LIST = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.GRID_MENU = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.BTN_ADD = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_DELETE = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_SAVE = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_REFRESH = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.FILE_LIST, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GRID_MENU.SuspendLayout()
        Me.SuspendLayout()
        '
        'FILE_LIST
        '
        Me.FILE_LIST.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.FILE_LIST.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FILE_LIST.Location = New System.Drawing.Point(0, 0)
        Me.FILE_LIST.Name = "FILE_LIST"
        Me.FILE_LIST.Rows.Count = 2
        Me.FILE_LIST.Rows.DefaultSize = 20
        Me.FILE_LIST.Size = New System.Drawing.Size(999, 632)
        Me.FILE_LIST.StyleInfo = resources.GetString("FILE_LIST.StyleInfo")
        Me.FILE_LIST.TabIndex = 2
        Me.FILE_LIST.VisualStyle = C1.Win.C1FlexGrid.VisualStyle.Office2010Blue
        '
        'GRID_MENU
        '
        Me.GRID_MENU.Font = New System.Drawing.Font("Arial", 9.0!)
        Me.GRID_MENU.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_ADD, Me.ToolStripSeparator1, Me.BTN_DELETE, Me.ToolStripSeparator2, Me.BTN_SAVE, Me.ToolStripSeparator3, Me.BTN_REFRESH})
        Me.GRID_MENU.Name = "GRID_MENU"
        Me.GRID_MENU.Size = New System.Drawing.Size(126, 110)
        '
        'BTN_ADD
        '
        'Me.BTN_ADD.Image = Global.YJ_MMS_Asset.My.Resources.Resources.plus
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
        'Me.BTN_DELETE.Image = Global.YJ_MMS_Asset.My.Resources.Resources.minus
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
        'Me.BTN_SAVE.Image = Global.YJ_MMS_Asset.My.Resources.Resources.save
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
        'PG_UPLOAD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(999, 632)
        Me.Controls.Add(Me.FILE_LIST)
        Me.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Name = "PG_UPLOAD"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "프로그램 파일 Upload"
        CType(Me.FILE_LIST, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GRID_MENU.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents FILE_LIST As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents GRID_MENU As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents BTN_ADD As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BTN_DELETE As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BTN_SAVE As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BTN_REFRESH As System.Windows.Forms.ToolStripMenuItem
End Class

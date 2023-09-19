<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Module_Out_PDA
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Module_Out_PDA))
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.BTN_Search = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Print = New System.Windows.Forms.ToolStripButton()
        Me.grid_OutList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.cms_Menu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.btn_RowDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Save2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.grid_OutList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cms_Menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_Search, Me.ToolStripSeparator1, Me.btn_Print})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(784, 25)
        Me.ToolStrip1.TabIndex = 3
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
        'btn_Print
        '
        Me.btn_Print.Enabled = False
        Me.btn_Print.Image = Global.Repair_System.My.Resources.Resources.print
        Me.btn_Print.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_Print.Name = "btn_Print"
        Me.btn_Print.Size = New System.Drawing.Size(63, 22)
        Me.btn_Print.Text = "프린트"
        '
        'grid_OutList
        '
        Me.grid_OutList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_OutList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_OutList.Location = New System.Drawing.Point(0, 25)
        Me.grid_OutList.Name = "grid_OutList"
        Me.grid_OutList.Rows.Count = 2
        Me.grid_OutList.Rows.DefaultSize = 20
        Me.grid_OutList.Size = New System.Drawing.Size(784, 348)
        Me.grid_OutList.StyleInfo = resources.GetString("grid_OutList.StyleInfo")
        Me.grid_OutList.TabIndex = 5
        '
        'cms_Menu
        '
        Me.cms_Menu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_RowDelete, Me.ToolStripSeparator4, Me.btn_Save2})
        Me.cms_Menu.Name = "GRID_MENU"
        Me.cms_Menu.Size = New System.Drawing.Size(115, 54)
        '
        'btn_RowDelete
        '
        Me.btn_RowDelete.Image = Global.Repair_System.My.Resources.Resources.minus_blue
        Me.btn_RowDelete.Name = "btn_RowDelete"
        Me.btn_RowDelete.Size = New System.Drawing.Size(114, 22)
        Me.btn_RowDelete.Text = "행 삭제"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(111, 6)
        '
        'btn_Save2
        '
        Me.btn_Save2.Image = Global.Repair_System.My.Resources.Resources.save2
        Me.btn_Save2.Name = "btn_Save2"
        Me.btn_Save2.Size = New System.Drawing.Size(114, 22)
        Me.btn_Save2.Text = "저장"
        '
        'frm_Module_Out_PDA
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 373)
        Me.Controls.Add(Me.grid_OutList)
        Me.Controls.Add(Me.ToolStrip1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_Module_Out_PDA"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "PDA 등록 리스트 확인"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        CType(Me.grid_OutList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cms_Menu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents BTN_Search As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents grid_OutList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents btn_Print As ToolStripButton
    Friend WithEvents cms_Menu As ContextMenuStrip
    Friend WithEvents btn_RowDelete As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents btn_Save2 As ToolStripMenuItem
End Class

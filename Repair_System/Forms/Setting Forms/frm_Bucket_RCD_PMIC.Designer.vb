<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Bucket_RCD_PMIC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Bucket_RCD_PMIC))
        Me.GRID_Data_List = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.btn_Search = New System.Windows.Forms.ToolStripButton()
        Me.btn_Menu_Save = New System.Windows.Forms.ToolStripButton()
        Me.cms_Menu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.btn_Row_Add = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_Row_Delete = New System.Windows.Forms.ToolStripMenuItem()
        Me.COL_VIEW = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_SAVE = New System.Windows.Forms.ToolStripMenuItem()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        CType(Me.GRID_Data_List, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.cms_Menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'GRID_Data_List
        '
        Me.GRID_Data_List.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.GRID_Data_List.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GRID_Data_List.Location = New System.Drawing.Point(0, 25)
        Me.GRID_Data_List.Name = "GRID_Data_List"
        Me.GRID_Data_List.Rows.Count = 2
        Me.GRID_Data_List.Rows.DefaultSize = 20
        Me.GRID_Data_List.Size = New System.Drawing.Size(1264, 748)
        Me.GRID_Data_List.StyleInfo = resources.GetString("GRID_Data_List.StyleInfo")
        Me.GRID_Data_List.TabIndex = 55
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Search, Me.btn_Menu_Save, Me.Form_CLose})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1264, 25)
        Me.ToolStrip1.TabIndex = 56
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'btn_Search
        '
        Me.btn_Search.Image = Global.Repair_System.My.Resources.Resources.Search
        Me.btn_Search.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_Search.Name = "btn_Search"
        Me.btn_Search.Size = New System.Drawing.Size(51, 22)
        Me.btn_Search.Text = "조회"
        '
        'btn_Menu_Save
        '
        Me.btn_Menu_Save.Image = Global.Repair_System.My.Resources.Resources.save4
        Me.btn_Menu_Save.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_Menu_Save.Name = "btn_Menu_Save"
        Me.btn_Menu_Save.Size = New System.Drawing.Size(51, 22)
        Me.btn_Menu_Save.Text = "저장"
        '
        'cms_Menu
        '
        Me.cms_Menu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Row_Add, Me.btn_Row_Delete, Me.COL_VIEW, Me.btn_SAVE})
        Me.cms_Menu.Name = "GRID_Menu"
        Me.cms_Menu.Size = New System.Drawing.Size(115, 76)
        '
        'btn_Row_Add
        '
        Me.btn_Row_Add.Name = "btn_Row_Add"
        Me.btn_Row_Add.Size = New System.Drawing.Size(114, 22)
        Me.btn_Row_Add.Text = "행 추가"
        '
        'btn_Row_Delete
        '
        Me.btn_Row_Delete.Name = "btn_Row_Delete"
        Me.btn_Row_Delete.Size = New System.Drawing.Size(114, 22)
        Me.btn_Row_Delete.Text = "행 삭제"
        '
        'COL_VIEW
        '
        Me.COL_VIEW.Name = "COL_VIEW"
        Me.COL_VIEW.Size = New System.Drawing.Size(111, 6)
        '
        'btn_SAVE
        '
        Me.btn_SAVE.Name = "btn_SAVE"
        Me.btn_SAVE.Size = New System.Drawing.Size(114, 22)
        Me.btn_SAVE.Text = "저장"
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
        'frm_Bucket_RCD_PMIC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.GRID_Data_List)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_Bucket_RCD_PMIC"
        Me.Text = "Bucket별 PMIC/RCD 조합 및 Vendor"
        CType(Me.GRID_Data_List, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.cms_Menu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GRID_Data_List As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents btn_Search As ToolStripButton
    Friend WithEvents btn_Menu_Save As ToolStripButton
    Friend WithEvents cms_Menu As ContextMenuStrip
    Friend WithEvents btn_Row_Add As ToolStripMenuItem
    Friend WithEvents btn_Row_Delete As ToolStripMenuItem
    Friend WithEvents COL_VIEW As ToolStripSeparator
    Friend WithEvents btn_SAVE As ToolStripMenuItem
    Friend WithEvents Form_CLose As ToolStripButton
End Class

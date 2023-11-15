<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CadData
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CadData))
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.BTN_CadData = New System.Windows.Forms.ToolStripButton()
        Me.Grid_CADData = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.BTN_FeederLoad = New System.Windows.Forms.ToolStripButton()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.Grid_CADData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_CadData, Me.BTN_FeederLoad})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(512, 25)
        Me.ToolStrip1.TabIndex = 2
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'BTN_CadData
        '
        Me.BTN_CadData.Image = CType(resources.GetObject("BTN_CadData.Image"), System.Drawing.Image)
        Me.BTN_CadData.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_CadData.Name = "BTN_CadData"
        Me.BTN_CadData.Size = New System.Drawing.Size(148, 22)
        Me.BTN_CadData.Text = "(Ref)xml파일 불러오기"
        '
        'Grid_CADData
        '
        Me.Grid_CADData.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.Grid_CADData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_CADData.Location = New System.Drawing.Point(0, 25)
        Me.Grid_CADData.Name = "Grid_CADData"
        Me.Grid_CADData.Rows.Count = 2
        Me.Grid_CADData.Rows.DefaultSize = 20
        Me.Grid_CADData.Size = New System.Drawing.Size(512, 644)
        Me.Grid_CADData.StyleInfo = resources.GetString("Grid_CADData.StyleInfo")
        Me.Grid_CADData.TabIndex = 5
        Me.Grid_CADData.VisualStyle = C1.Win.C1FlexGrid.VisualStyle.Office2010Blue
        '
        'BTN_FeederLoad
        '
        Me.BTN_FeederLoad.Image = CType(resources.GetObject("BTN_FeederLoad.Image"), System.Drawing.Image)
        Me.BTN_FeederLoad.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_FeederLoad.Name = "BTN_FeederLoad"
        Me.BTN_FeederLoad.Size = New System.Drawing.Size(166, 22)
        Me.BTN_FeederLoad.Text = "(Feeder)xml파일 불러오기"
        '
        'CadData
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(512, 669)
        Me.Controls.Add(Me.Grid_CADData)
        Me.Controls.Add(Me.ToolStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CadData"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SAMSUNG CAD Data"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        CType(Me.Grid_CADData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents BTN_CadData As ToolStripButton
    Friend WithEvents Grid_CADData As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents BTN_FeederLoad As ToolStripButton
End Class

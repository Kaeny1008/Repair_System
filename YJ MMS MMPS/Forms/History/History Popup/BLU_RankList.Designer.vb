<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BLU_RankList
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BLU_RankList))
        Me.GRID_RankList = New C1.Win.C1FlexGrid.C1FlexGrid()
        CType(Me.GRID_RankList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GRID_RankList
        '
        Me.GRID_RankList.AllowFreezing = C1.Win.C1FlexGrid.AllowFreezingEnum.Both
        Me.GRID_RankList.AutoResize = True
        Me.GRID_RankList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.GRID_RankList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GRID_RankList.Font = New System.Drawing.Font("돋움", 9.0!)
        Me.GRID_RankList.KeyActionTab = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross
        Me.GRID_RankList.Location = New System.Drawing.Point(0, 0)
        Me.GRID_RankList.Name = "GRID_RankList"
        Me.GRID_RankList.Rows.Count = 2
        Me.GRID_RankList.Rows.DefaultSize = 20
        Me.GRID_RankList.Size = New System.Drawing.Size(211, 224)
        Me.GRID_RankList.StyleInfo = resources.GetString("GRID_RankList.StyleInfo")
        Me.GRID_RankList.TabIndex = 1
        Me.GRID_RankList.VisualStyle = C1.Win.C1FlexGrid.VisualStyle.Office2010Blue
        '
        'BLU_RankList
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(211, 224)
        Me.Controls.Add(Me.GRID_RankList)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BLU_RankList"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "BLU Rank List"
        CType(Me.GRID_RankList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GRID_RankList As C1.Win.C1FlexGrid.C1FlexGrid
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AlarmHistory
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AlarmHistory))
        Me.Grid_historyList = New C1.Win.C1FlexGrid.C1FlexGrid()
        CType(Me.Grid_historyList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Grid_historyList
        '
        Me.Grid_historyList.AllowFreezing = C1.Win.C1FlexGrid.AllowFreezingEnum.Both
        Me.Grid_historyList.AutoResize = True
        Me.Grid_historyList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.Grid_historyList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_historyList.Font = New System.Drawing.Font("돋움", 9.0!)
        Me.Grid_historyList.KeyActionTab = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross
        Me.Grid_historyList.Location = New System.Drawing.Point(0, 0)
        Me.Grid_historyList.Name = "Grid_historyList"
        Me.Grid_historyList.Rows.Count = 2
        Me.Grid_historyList.Rows.DefaultSize = 20
        Me.Grid_historyList.Size = New System.Drawing.Size(903, 561)
        Me.Grid_historyList.StyleInfo = resources.GetString("Grid_historyList.StyleInfo")
        Me.Grid_historyList.TabIndex = 1
        Me.Grid_historyList.VisualStyle = C1.Win.C1FlexGrid.VisualStyle.Office2010Blue
        '
        'AlarmHistory
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(903, 561)
        Me.Controls.Add(Me.Grid_historyList)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AlarmHistory"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "알람이력 확인"
        CType(Me.Grid_historyList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Grid_historyList As C1.Win.C1FlexGrid.C1FlexGrid
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Material_Use_History
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Material_Use_History))
        Me.grid_Lot_List = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.grid_Lot_List, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grid_Lot_List
        '
        Me.grid_Lot_List.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_Lot_List.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_Lot_List.Location = New System.Drawing.Point(0, 0)
        Me.grid_Lot_List.Name = "grid_Lot_List"
        Me.grid_Lot_List.Rows.Count = 2
        Me.grid_Lot_List.Rows.DefaultSize = 20
        Me.grid_Lot_List.Size = New System.Drawing.Size(1166, 635)
        Me.grid_Lot_List.StyleInfo = resources.GetString("grid_Lot_List.StyleInfo")
        Me.grid_Lot_List.TabIndex = 1
        '
        'frm_Material_Use_History
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1166, 635)
        Me.Controls.Add(Me.grid_Lot_List)
        Me.Name = "frm_Material_Use_History"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frm_Material_Use_History"
        CType(Me.grid_Lot_List, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents grid_Lot_List As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents ToolTip1 As ToolTip
End Class

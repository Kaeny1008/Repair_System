<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ModelSearch
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ModelSearch))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.BTN_ModelSearch = New System.Windows.Forms.Button()
        Me.TB_ModelName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GRID_ModelList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.GRID_ModelList, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BTN_ModelSearch)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TB_ModelName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GRID_ModelList)
        Me.SplitContainer1.Size = New System.Drawing.Size(709, 415)
        Me.SplitContainer1.SplitterDistance = 42
        Me.SplitContainer1.TabIndex = 0
        '
        'BTN_ModelSearch
        '
        Me.BTN_ModelSearch.Location = New System.Drawing.Point(659, 19)
        Me.BTN_ModelSearch.Margin = New System.Windows.Forms.Padding(0)
        Me.BTN_ModelSearch.Name = "BTN_ModelSearch"
        Me.BTN_ModelSearch.Size = New System.Drawing.Size(50, 23)
        Me.BTN_ModelSearch.TabIndex = 125
        Me.BTN_ModelSearch.Text = "검색"
        Me.BTN_ModelSearch.UseVisualStyleBackColor = True
        '
        'TB_ModelName
        '
        Me.TB_ModelName.BackColor = System.Drawing.Color.White
        Me.TB_ModelName.Location = New System.Drawing.Point(150, 20)
        Me.TB_ModelName.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_ModelName.Name = "TB_ModelName"
        Me.TB_ModelName.Size = New System.Drawing.Size(509, 21)
        Me.TB_ModelName.TabIndex = 124
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(0, 20)
        Me.Label4.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(150, 21)
        Me.Label4.TabIndex = 123
        Me.Label4.Text = "검색어"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GRID_ModelList
        '
        Me.GRID_ModelList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.GRID_ModelList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GRID_ModelList.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.GRID_ModelList.Location = New System.Drawing.Point(0, 0)
        Me.GRID_ModelList.Name = "GRID_ModelList"
        Me.GRID_ModelList.Rows.Count = 2
        Me.GRID_ModelList.Rows.DefaultSize = 20
        Me.GRID_ModelList.Size = New System.Drawing.Size(709, 369)
        Me.GRID_ModelList.StyleInfo = resources.GetString("GRID_ModelList.StyleInfo")
        Me.GRID_ModelList.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(148, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(269, 12)
        Me.Label1.TabIndex = 126
        Me.Label1.Text = "※ 영문 대소문자를 구분하여 검색하여 주십시오."
        '
        'ModelSearch
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(709, 415)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ModelSearch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Model Search"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.GRID_ModelList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TB_ModelName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BTN_ModelSearch As System.Windows.Forms.Button
    Friend WithEvents GRID_ModelList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Label1 As Label
End Class

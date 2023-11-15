<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HistoryBLUForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(HistoryBLUForm))
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.TB_Marking = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Cb_result = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Cb_workLine2 = New System.Windows.Forms.ComboBox()
        Me.Tb_modelName2 = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.DTP_endDate2 = New System.Windows.Forms.DateTimePicker()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.DTP_startDate2 = New System.Windows.Forms.DateTimePicker()
        Me.Btn_changeSearch = New System.Windows.Forms.Button()
        Me.Grid_PartsChange = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.Grid_PartsChange, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TB_Marking)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label15)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Cb_result)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Cb_workLine2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Tb_modelName2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DTP_endDate2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label13)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label14)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DTP_startDate2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Btn_changeSearch)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Grid_PartsChange)
        Me.SplitContainer2.Size = New System.Drawing.Size(1264, 773)
        Me.SplitContainer2.SplitterDistance = 64
        Me.SplitContainer2.TabIndex = 3
        '
        'TB_Marking
        '
        Me.TB_Marking.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Marking.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Marking.Location = New System.Drawing.Point(632, 33)
        Me.TB_Marking.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Marking.Name = "TB_Marking"
        Me.TB_Marking.Size = New System.Drawing.Size(168, 21)
        Me.TB_Marking.TabIndex = 32
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(505, 33)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(127, 21)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "Marking or Lot No."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label15.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(800, 10)
        Me.Label15.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(127, 21)
        Me.Label15.TabIndex = 29
        Me.Label15.Text = "결과"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_result
        '
        Me.Cb_result.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_result.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_result.FormattingEnabled = True
        Me.Cb_result.Items.AddRange(New Object() {"ALL", "OK", "NG"})
        Me.Cb_result.Location = New System.Drawing.Point(927, 10)
        Me.Cb_result.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_result.Name = "Cb_result"
        Me.Cb_result.Size = New System.Drawing.Size(168, 21)
        Me.Cb_result.TabIndex = 30
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(505, 10)
        Me.Label6.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(127, 21)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "작업 Line"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Cb_workLine2
        '
        Me.Cb_workLine2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_workLine2.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.Cb_workLine2.FormattingEnabled = True
        Me.Cb_workLine2.Items.AddRange(New Object() {"A-LINE", "B-LINE"})
        Me.Cb_workLine2.Location = New System.Drawing.Point(632, 10)
        Me.Cb_workLine2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Cb_workLine2.Name = "Cb_workLine2"
        Me.Cb_workLine2.Size = New System.Drawing.Size(168, 21)
        Me.Cb_workLine2.TabIndex = 26
        '
        'Tb_modelName2
        '
        Me.Tb_modelName2.BackColor = System.Drawing.SystemColors.Window
        Me.Tb_modelName2.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Tb_modelName2.Location = New System.Drawing.Point(246, 33)
        Me.Tb_modelName2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_modelName2.Name = "Tb_modelName2"
        Me.Tb_modelName2.Size = New System.Drawing.Size(259, 21)
        Me.Tb_modelName2.TabIndex = 22
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label12.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.Label12.ForeColor = System.Drawing.Color.White
        Me.Label12.Location = New System.Drawing.Point(119, 33)
        Me.Label12.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(127, 21)
        Me.Label12.TabIndex = 21
        Me.Label12.Text = "모델명"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DTP_endDate2
        '
        Me.DTP_endDate2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTP_endDate2.Location = New System.Drawing.Point(384, 10)
        Me.DTP_endDate2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DTP_endDate2.Name = "DTP_endDate2"
        Me.DTP_endDate2.Size = New System.Drawing.Size(121, 21)
        Me.DTP_endDate2.TabIndex = 18
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(370, 14)
        Me.Label13.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(14, 12)
        Me.Label13.TabIndex = 17
        Me.Label13.Text = "~"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.RoyalBlue
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label14.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(119, 10)
        Me.Label14.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(127, 21)
        Me.Label14.TabIndex = 15
        Me.Label14.Text = "조회기간"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DTP_startDate2
        '
        Me.DTP_startDate2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTP_startDate2.Location = New System.Drawing.Point(246, 10)
        Me.DTP_startDate2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DTP_startDate2.Name = "DTP_startDate2"
        Me.DTP_startDate2.Size = New System.Drawing.Size(121, 21)
        Me.DTP_startDate2.TabIndex = 16
        '
        'Btn_changeSearch
        '
        Me.Btn_changeSearch.Font = New System.Drawing.Font("굴림", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_changeSearch.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.Search
        Me.Btn_changeSearch.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Btn_changeSearch.Location = New System.Drawing.Point(11, 12)
        Me.Btn_changeSearch.Name = "Btn_changeSearch"
        Me.Btn_changeSearch.Size = New System.Drawing.Size(96, 42)
        Me.Btn_changeSearch.TabIndex = 8
        Me.Btn_changeSearch.Text = "조회"
        Me.Btn_changeSearch.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Btn_changeSearch.UseVisualStyleBackColor = True
        '
        'Grid_PartsChange
        '
        Me.Grid_PartsChange.AllowFreezing = C1.Win.C1FlexGrid.AllowFreezingEnum.Both
        Me.Grid_PartsChange.AutoResize = True
        Me.Grid_PartsChange.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.Grid_PartsChange.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid_PartsChange.Font = New System.Drawing.Font("돋움", 9.0!)
        Me.Grid_PartsChange.KeyActionTab = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross
        Me.Grid_PartsChange.Location = New System.Drawing.Point(0, 0)
        Me.Grid_PartsChange.Name = "Grid_PartsChange"
        Me.Grid_PartsChange.Rows.Count = 2
        Me.Grid_PartsChange.Rows.DefaultSize = 20
        Me.Grid_PartsChange.Size = New System.Drawing.Size(1264, 705)
        Me.Grid_PartsChange.StyleInfo = resources.GetString("Grid_PartsChange.StyleInfo")
        Me.Grid_PartsChange.TabIndex = 55
        Me.Grid_PartsChange.VisualStyle = C1.Win.C1FlexGrid.VisualStyle.Office2010Blue
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(803, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(524, 12)
        Me.Label2.TabIndex = 33
        Me.Label2.Text = "※ 자재구분이 'LED'항목을 더블클릭 하시면 당시 사용가능한 Rank 목록을 확인 할 수 있습니다."
        '
        'HistoryBLUForm
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.SplitContainer2)
        Me.Name = "HistoryBLUForm"
        Me.Text = "이력보기(BLU)"
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.Grid_PartsChange, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents Label15 As Label
    Friend WithEvents Cb_result As ComboBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Cb_workLine2 As ComboBox
    Friend WithEvents Tb_modelName2 As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents DTP_endDate2 As DateTimePicker
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents DTP_startDate2 As DateTimePicker
    Friend WithEvents Btn_changeSearch As Button
    Friend WithEvents Grid_PartsChange As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents TB_Marking As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
End Class

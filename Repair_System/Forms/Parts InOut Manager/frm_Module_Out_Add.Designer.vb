<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_Module_Out_Add
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Module_Out_Add))
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.BTN_Search = New System.Windows.Forms.ToolStripButton()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_PDA_Check = New System.Windows.Forms.ToolStripButton()
        Me.btn_New_Kitting = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Save = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.grid_OutList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TB_S_YJ_No = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.grid_ShipList = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btn_PDA_TempLoad = New System.Windows.Forms.Button()
        Me.btn_TempLoad = New System.Windows.Forms.Button()
        Me.btn_TempSave = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.tb_SlipNo = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ToolStrip1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grid_OutList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.grid_ShipList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_Search, Me.Form_CLose, Me.ToolStripSeparator1, Me.btn_PDA_Check, Me.btn_New_Kitting, Me.ToolStripSeparator2, Me.btn_Save, Me.ToolStripSeparator3})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1264, 25)
        Me.ToolStrip1.TabIndex = 2
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
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'btn_PDA_Check
        '
        Me.btn_PDA_Check.Image = Global.Repair_System.My.Resources.Resources.phone_settings_128
        Me.btn_PDA_Check.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_PDA_Check.Name = "btn_PDA_Check"
        Me.btn_PDA_Check.Size = New System.Drawing.Size(143, 22)
        Me.btn_PDA_Check.Text = "PDA 등록리스트 확인"
        '
        'btn_New_Kitting
        '
        Me.btn_New_Kitting.Image = Global.Repair_System.My.Resources.Resources.TEST_FOLDER
        Me.btn_New_Kitting.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_New_Kitting.Name = "btn_New_Kitting"
        Me.btn_New_Kitting.Size = New System.Drawing.Size(75, 22)
        Me.btn_New_Kitting.Text = "신규등록"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'btn_Save
        '
        Me.btn_Save.Enabled = False
        Me.btn_Save.Image = Global.Repair_System.My.Resources.Resources.save2
        Me.btn_Save.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_Save.Name = "btn_Save"
        Me.btn_Save.Size = New System.Drawing.Size(51, 22)
        Me.btn_Save.Text = "저장"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.grid_OutList)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.grid_ShipList)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 748)
        Me.SplitContainer1.SplitterDistance = 421
        Me.SplitContainer1.TabIndex = 3
        '
        'grid_OutList
        '
        Me.grid_OutList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_OutList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_OutList.Location = New System.Drawing.Point(0, 65)
        Me.grid_OutList.Name = "grid_OutList"
        Me.grid_OutList.Rows.Count = 2
        Me.grid_OutList.Rows.DefaultSize = 20
        Me.grid_OutList.Size = New System.Drawing.Size(421, 683)
        Me.grid_OutList.StyleInfo = resources.GetString("grid_OutList.StyleInfo")
        Me.grid_OutList.TabIndex = 4
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.TB_S_YJ_No)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.DateTimePicker2)
        Me.Panel1.Controls.Add(Me.DateTimePicker1)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(421, 65)
        Me.Panel1.TabIndex = 5
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(362, 18)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(42, 32)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "Test"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'TB_S_YJ_No
        '
        Me.TB_S_YJ_No.BackColor = System.Drawing.SystemColors.Window
        Me.TB_S_YJ_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_S_YJ_No.Location = New System.Drawing.Point(87, 31)
        Me.TB_S_YJ_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_S_YJ_No.Name = "TB_S_YJ_No"
        Me.TB_S_YJ_No.Size = New System.Drawing.Size(242, 21)
        Me.TB_S_YJ_No.TabIndex = 8
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(9, 31)
        Me.Label8.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(78, 21)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "Slip No."
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(193, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(30, 21)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "~"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker2.Location = New System.Drawing.Point(226, 8)
        Me.DateTimePicker2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(103, 21)
        Me.DateTimePicker2.TabIndex = 3
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(87, 8)
        Me.DateTimePicker1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(103, 21)
        Me.DateTimePicker1.TabIndex = 1
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label11.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.White
        Me.Label11.Location = New System.Drawing.Point(9, 8)
        Me.Label11.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(78, 21)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "출고일자"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'grid_ShipList
        '
        Me.grid_ShipList.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_ShipList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_ShipList.Location = New System.Drawing.Point(0, 96)
        Me.grid_ShipList.Name = "grid_ShipList"
        Me.grid_ShipList.Rows.Count = 2
        Me.grid_ShipList.Rows.DefaultSize = 20
        Me.grid_ShipList.Size = New System.Drawing.Size(839, 652)
        Me.grid_ShipList.StyleInfo = resources.GetString("grid_ShipList.StyleInfo")
        Me.grid_ShipList.TabIndex = 10
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Panel3.Controls.Add(Me.btn_PDA_TempLoad)
        Me.Panel3.Controls.Add(Me.btn_TempLoad)
        Me.Panel3.Controls.Add(Me.btn_TempSave)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.CheckBox1)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.Label10)
        Me.Panel3.Controls.Add(Me.tb_SlipNo)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(839, 96)
        Me.Panel3.TabIndex = 9
        '
        'btn_PDA_TempLoad
        '
        Me.btn_PDA_TempLoad.Enabled = False
        Me.btn_PDA_TempLoad.Location = New System.Drawing.Point(314, 69)
        Me.btn_PDA_TempLoad.Name = "btn_PDA_TempLoad"
        Me.btn_PDA_TempLoad.Size = New System.Drawing.Size(192, 20)
        Me.btn_PDA_TempLoad.TabIndex = 20
        Me.btn_PDA_TempLoad.Text = "PDA 데이터 불러오기"
        Me.btn_PDA_TempLoad.UseVisualStyleBackColor = True
        '
        'btn_TempLoad
        '
        Me.btn_TempLoad.Enabled = False
        Me.btn_TempLoad.Location = New System.Drawing.Point(314, 50)
        Me.btn_TempLoad.Name = "btn_TempLoad"
        Me.btn_TempLoad.Size = New System.Drawing.Size(192, 20)
        Me.btn_TempLoad.TabIndex = 19
        Me.btn_TempLoad.Text = "임시저장 데이터 불러오기"
        Me.btn_TempLoad.UseVisualStyleBackColor = True
        '
        'btn_TempSave
        '
        Me.btn_TempSave.Enabled = False
        Me.btn_TempSave.Location = New System.Drawing.Point(251, 50)
        Me.btn_TempSave.Name = "btn_TempSave"
        Me.btn_TempSave.Size = New System.Drawing.Size(68, 20)
        Me.btn_TempSave.TabIndex = 18
        Me.btn_TempSave.Text = "임시저장"
        Me.btn_TempSave.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(629, 73)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(201, 16)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "총 출고 Module : "
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(15, 58)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(189, 16)
        Me.CheckBox1.TabIndex = 16
        Me.CheckBox1.Text = "강제출고(YJ No.없는 Lot No.)"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.DimGray
        Me.Label1.Location = New System.Drawing.Point(13, 77)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(163, 12)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "( F1 ) 행 추가, ( F3 ) 행 삭제"
        '
        'Label10
        '
        Me.Label10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label10.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(629, 53)
        Me.Label10.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(201, 16)
        Me.Label10.TabIndex = 12
        Me.Label10.Text = "총 출고 Lot수 : "
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_SlipNo
        '
        Me.tb_SlipNo.Enabled = False
        Me.tb_SlipNo.Font = New System.Drawing.Font("굴림", 12.0!)
        Me.tb_SlipNo.Location = New System.Drawing.Point(162, 18)
        Me.tb_SlipNo.Margin = New System.Windows.Forms.Padding(0, 3, 3, 3)
        Me.tb_SlipNo.Name = "tb_SlipNo"
        Me.tb_SlipNo.Size = New System.Drawing.Size(315, 26)
        Me.tb_SlipNo.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(15, 18)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(147, 26)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Slip No."
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frm_Module_Out_Add
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_Module_Out_Add"
        Me.Text = "출고 등록(전표생성)"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grid_OutList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.grid_ShipList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents BTN_Search As ToolStripButton
    Friend WithEvents Form_CLose As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents btn_New_Kitting As ToolStripButton
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents btn_Save As ToolStripButton
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents grid_OutList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents TB_S_YJ_No As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents DateTimePicker2 As DateTimePicker
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents Label11 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents grid_ShipList As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Label10 As Label
    Friend WithEvents tb_SlipNo As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Label3 As Label
    Friend WithEvents btn_TempLoad As Button
    Friend WithEvents btn_TempSave As Button
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents btn_PDA_Check As ToolStripButton
    Friend WithEvents btn_PDA_TempLoad As Button
End Class

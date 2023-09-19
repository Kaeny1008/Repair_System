<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Material_In_Add
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Material_In_Add))
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.Delete_Cancel = New System.Windows.Forms.ToolStripMenuItem()
        Me.cms_GridMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.Row_ADD = New System.Windows.Forms.ToolStripMenuItem()
        Me.Row_Delete = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_Save2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.BTN_Search = New System.Windows.Forms.ToolStripButton()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        Me.btn_NewParts = New System.Windows.Forms.ToolStripButton()
        Me.BTN_Save = New System.Windows.Forms.ToolStripButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.CB_Parts_Section = New System.Windows.Forms.ComboBox()
        Me.TB_Slip_No = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dtp_End = New System.Windows.Forms.DateTimePicker()
        Me.dtp_Start = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.grid_SlipNo = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.btn_LoadStop = New System.Windows.Forms.Button()
        Me.btn_File_Load = New System.Windows.Forms.Button()
        Me.grid_Parts_List = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.cms_GridMenu.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.grid_SlipNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grid_Parts_List, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(123, 6)
        '
        'Delete_Cancel
        '
        Me.Delete_Cancel.Image = Global.Repair_System.My.Resources.Resources.arrow_double
        Me.Delete_Cancel.Name = "Delete_Cancel"
        Me.Delete_Cancel.Size = New System.Drawing.Size(126, 22)
        Me.Delete_Cancel.Text = "삭제 취소"
        '
        'cms_GridMenu
        '
        Me.cms_GridMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Row_ADD, Me.Row_Delete, Me.ToolStripSeparator3, Me.Delete_Cancel, Me.ToolStripSeparator4, Me.BTN_Save2})
        Me.cms_GridMenu.Name = "GRID_MENU"
        Me.cms_GridMenu.Size = New System.Drawing.Size(127, 104)
        '
        'Row_ADD
        '
        Me.Row_ADD.Image = Global.Repair_System.My.Resources.Resources.plus_blue
        Me.Row_ADD.Name = "Row_ADD"
        Me.Row_ADD.Size = New System.Drawing.Size(126, 22)
        Me.Row_ADD.Text = "행 추가"
        '
        'Row_Delete
        '
        Me.Row_Delete.Image = Global.Repair_System.My.Resources.Resources.minus_blue
        Me.Row_Delete.Name = "Row_Delete"
        Me.Row_Delete.Size = New System.Drawing.Size(126, 22)
        Me.Row_Delete.Text = "행 삭제"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(123, 6)
        '
        'BTN_Save2
        '
        Me.BTN_Save2.Image = Global.Repair_System.My.Resources.Resources.save2
        Me.BTN_Save2.Name = "BTN_Save2"
        Me.BTN_Save2.Size = New System.Drawing.Size(126, 22)
        Me.BTN_Save2.Text = "저장"
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
        'btn_NewParts
        '
        Me.btn_NewParts.Image = Global.Repair_System.My.Resources.Resources.TEST_FOLDER
        Me.btn_NewParts.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_NewParts.Name = "btn_NewParts"
        Me.btn_NewParts.Size = New System.Drawing.Size(75, 22)
        Me.btn_NewParts.Text = "신규등록"
        '
        'BTN_Save
        '
        Me.BTN_Save.Enabled = False
        Me.BTN_Save.Image = Global.Repair_System.My.Resources.Resources.save2
        Me.BTN_Save.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_Save.Name = "BTN_Save"
        Me.BTN_Save.Size = New System.Drawing.Size(51, 22)
        Me.BTN_Save.Text = "저장"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Panel1.Controls.Add(Me.CB_Parts_Section)
        Me.Panel1.Controls.Add(Me.TB_Slip_No)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.dtp_End)
        Me.Panel1.Controls.Add(Me.dtp_Start)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(360, 85)
        Me.Panel1.TabIndex = 0
        '
        'CB_Parts_Section
        '
        Me.CB_Parts_Section.Enabled = False
        Me.CB_Parts_Section.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.CB_Parts_Section.FormattingEnabled = True
        Me.CB_Parts_Section.Items.AddRange(New Object() {"", "Module", "Component"})
        Me.CB_Parts_Section.Location = New System.Drawing.Point(86, 31)
        Me.CB_Parts_Section.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.CB_Parts_Section.Name = "CB_Parts_Section"
        Me.CB_Parts_Section.Size = New System.Drawing.Size(242, 21)
        Me.CB_Parts_Section.TabIndex = 5
        '
        'TB_Slip_No
        '
        Me.TB_Slip_No.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Slip_No.Enabled = False
        Me.TB_Slip_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Slip_No.Location = New System.Drawing.Point(86, 54)
        Me.TB_Slip_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Slip_No.Name = "TB_Slip_No"
        Me.TB_Slip_No.Size = New System.Drawing.Size(242, 21)
        Me.TB_Slip_No.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(8, 54)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 21)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Slip No."
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(8, 31)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 21)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "구분"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(192, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 21)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "~"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dtp_End
        '
        Me.dtp_End.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.dtp_End.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp_End.Location = New System.Drawing.Point(225, 8)
        Me.dtp_End.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.dtp_End.Name = "dtp_End"
        Me.dtp_End.Size = New System.Drawing.Size(103, 21)
        Me.dtp_End.TabIndex = 3
        '
        'dtp_Start
        '
        Me.dtp_Start.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.dtp_Start.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp_Start.Location = New System.Drawing.Point(86, 8)
        Me.dtp_Start.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.dtp_Start.Name = "dtp_Start"
        Me.dtp_Start.Size = New System.Drawing.Size(103, 21)
        Me.dtp_Start.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 21)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "입고 일자"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'grid_SlipNo
        '
        Me.grid_SlipNo.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_SlipNo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_SlipNo.Location = New System.Drawing.Point(0, 85)
        Me.grid_SlipNo.Name = "grid_SlipNo"
        Me.grid_SlipNo.Rows.Count = 2
        Me.grid_SlipNo.Rows.DefaultSize = 20
        Me.grid_SlipNo.Size = New System.Drawing.Size(360, 663)
        Me.grid_SlipNo.StyleInfo = resources.GetString("grid_SlipNo.StyleInfo")
        Me.grid_SlipNo.TabIndex = 1
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.grid_SlipNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btn_LoadStop)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btn_File_Load)
        Me.SplitContainer1.Panel2.Controls.Add(Me.grid_Parts_List)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 748)
        Me.SplitContainer1.SplitterDistance = 360
        Me.SplitContainer1.TabIndex = 3
        '
        'btn_LoadStop
        '
        Me.btn_LoadStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_LoadStop.Image = Global.Repair_System.My.Resources.Resources.cancel
        Me.btn_LoadStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_LoadStop.Location = New System.Drawing.Point(636, 8)
        Me.btn_LoadStop.Name = "btn_LoadStop"
        Me.btn_LoadStop.Size = New System.Drawing.Size(126, 34)
        Me.btn_LoadStop.TabIndex = 7
        Me.btn_LoadStop.Text = "불러오기 중지"
        Me.btn_LoadStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btn_LoadStop.UseVisualStyleBackColor = False
        '
        'btn_File_Load
        '
        Me.btn_File_Load.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_File_Load.Enabled = False
        Me.btn_File_Load.Location = New System.Drawing.Point(768, 8)
        Me.btn_File_Load.Name = "btn_File_Load"
        Me.btn_File_Load.Size = New System.Drawing.Size(120, 34)
        Me.btn_File_Load.TabIndex = 6
        Me.btn_File_Load.Text = "입고 List 불러오기"
        Me.btn_File_Load.UseVisualStyleBackColor = True
        '
        'grid_Parts_List
        '
        Me.grid_Parts_List.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_Parts_List.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_Parts_List.Location = New System.Drawing.Point(0, 48)
        Me.grid_Parts_List.Name = "grid_Parts_List"
        Me.grid_Parts_List.Rows.Count = 2
        Me.grid_Parts_List.Rows.DefaultSize = 20
        Me.grid_Parts_List.Size = New System.Drawing.Size(900, 700)
        Me.grid_Parts_List.StyleInfo = resources.GetString("grid_Parts_List.StyleInfo")
        Me.grid_Parts_List.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label4.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold)
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(0, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(900, 48)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "입고 List"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_Search, Me.ToolStripSeparator1, Me.btn_NewParts, Me.ToolStripSeparator2, Me.BTN_Save, Me.Form_CLose, Me.ToolStripSeparator6})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1264, 25)
        Me.ToolStrip1.TabIndex = 2
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(6, 25)
        '
        'frm_Material_In_Add
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_Material_In_Add"
        Me.Text = "자재 입고 등록"
        Me.cms_GridMenu.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.grid_SlipNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grid_Parts_List, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Delete_Cancel As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cms_GridMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents Row_ADD As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Row_Delete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BTN_Save2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BTN_Search As System.Windows.Forms.ToolStripButton
    Friend WithEvents Form_CLose As System.Windows.Forms.ToolStripButton
    Friend WithEvents btn_NewParts As System.Windows.Forms.ToolStripButton
    Friend WithEvents BTN_Save As System.Windows.Forms.ToolStripButton
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents CB_Parts_Section As System.Windows.Forms.ComboBox
    Friend WithEvents TB_Slip_No As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtp_End As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtp_Start As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grid_SlipNo As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents grid_Parts_List As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Label4 As Label
    Friend WithEvents btn_LoadStop As Button
    Friend WithEvents btn_File_Load As Button
End Class

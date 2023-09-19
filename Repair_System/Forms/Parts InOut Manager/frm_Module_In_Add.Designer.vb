<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_Module_In_Add
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Module_In_Add))
        Me.tb_Slip_No = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dtp_End = New System.Windows.Forms.DateTimePicker()
        Me.BTN_File_Load1 = New System.Windows.Forms.Button()
        Me.cms_Menu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.btn_RowAdd = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_RowDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Multi_Row_Insert = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Delete_Cancel = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Save2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtp_Start = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.grid_SlipNo = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.grid_SlipNoDetail = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.btn_Search = New System.Windows.Forms.ToolStripButton()
        Me.btn_New_SlipNo = New System.Windows.Forms.ToolStripButton()
        Me.btn_Save = New System.Windows.Forms.ToolStripButton()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.BTN_File_Load2 = New System.Windows.Forms.Button()
        Me.BTN_File_Load3 = New System.Windows.Forms.Button()
        Me.cms_Menu.SuspendLayout()
        CType(Me.grid_SlipNo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grid_SlipNoDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tb_Slip_No
        '
        Me.tb_Slip_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Slip_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Slip_No.Location = New System.Drawing.Point(86, 31)
        Me.tb_Slip_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Slip_No.Name = "tb_Slip_No"
        Me.tb_Slip_No.Size = New System.Drawing.Size(242, 21)
        Me.tb_Slip_No.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(8, 31)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 21)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Slip No."
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
        'BTN_File_Load1
        '
        Me.BTN_File_Load1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BTN_File_Load1.Enabled = False
        Me.BTN_File_Load1.Location = New System.Drawing.Point(516, 8)
        Me.BTN_File_Load1.Name = "BTN_File_Load1"
        Me.BTN_File_Load1.Size = New System.Drawing.Size(120, 34)
        Me.BTN_File_Load1.TabIndex = 2
        Me.BTN_File_Load1.Text = "1. 입고 List" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "불러오기 (OMS)"
        Me.BTN_File_Load1.UseVisualStyleBackColor = True
        '
        'cms_Menu
        '
        Me.cms_Menu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_RowAdd, Me.btn_RowDelete, Me.ToolStripSeparator3, Me.btn_Multi_Row_Insert, Me.ToolStripSeparator5, Me.btn_Delete_Cancel, Me.ToolStripSeparator4, Me.btn_Save2})
        Me.cms_Menu.Name = "GRID_MENU"
        Me.cms_Menu.Size = New System.Drawing.Size(143, 132)
        '
        'btn_RowAdd
        '
        Me.btn_RowAdd.Image = Global.Repair_System.My.Resources.Resources.plus_blue
        Me.btn_RowAdd.Name = "btn_RowAdd"
        Me.btn_RowAdd.Size = New System.Drawing.Size(142, 22)
        Me.btn_RowAdd.Text = "행 추가"
        '
        'btn_RowDelete
        '
        Me.btn_RowDelete.Image = Global.Repair_System.My.Resources.Resources.minus_blue
        Me.btn_RowDelete.Name = "btn_RowDelete"
        Me.btn_RowDelete.Size = New System.Drawing.Size(142, 22)
        Me.btn_RowDelete.Text = "행 삭제"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(139, 6)
        '
        'btn_Multi_Row_Insert
        '
        Me.btn_Multi_Row_Insert.Enabled = False
        Me.btn_Multi_Row_Insert.Name = "btn_Multi_Row_Insert"
        Me.btn_Multi_Row_Insert.Size = New System.Drawing.Size(142, 22)
        Me.btn_Multi_Row_Insert.Text = "다중 행 추가"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(139, 6)
        '
        'btn_Delete_Cancel
        '
        Me.btn_Delete_Cancel.Image = Global.Repair_System.My.Resources.Resources.arrow_double
        Me.btn_Delete_Cancel.Name = "btn_Delete_Cancel"
        Me.btn_Delete_Cancel.Size = New System.Drawing.Size(142, 22)
        Me.btn_Delete_Cancel.Text = "삭제 취소"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(139, 6)
        '
        'btn_Save2
        '
        Me.btn_Save2.Image = Global.Repair_System.My.Resources.Resources.save2
        Me.btn_Save2.Name = "btn_Save2"
        Me.btn_Save2.Size = New System.Drawing.Size(142, 22)
        Me.btn_Save2.Text = "저장"
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
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "입고 List"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomLeft
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
        Me.Label1.Text = "저장 일자"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'grid_SlipNo
        '
        Me.grid_SlipNo.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_SlipNo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_SlipNo.Location = New System.Drawing.Point(0, 59)
        Me.grid_SlipNo.Name = "grid_SlipNo"
        Me.grid_SlipNo.Rows.Count = 2
        Me.grid_SlipNo.Rows.DefaultSize = 20
        Me.grid_SlipNo.Size = New System.Drawing.Size(360, 689)
        Me.grid_SlipNo.StyleInfo = resources.GetString("grid_SlipNo.StyleInfo")
        Me.grid_SlipNo.TabIndex = 1
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Panel1.Controls.Add(Me.tb_Slip_No)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.dtp_End)
        Me.Panel1.Controls.Add(Me.dtp_Start)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(360, 59)
        Me.Panel1.TabIndex = 0
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.grid_SlipNoDetail)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BTN_File_Load3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BTN_File_Load2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BTN_File_Load1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 748)
        Me.SplitContainer1.SplitterDistance = 360
        Me.SplitContainer1.TabIndex = 3
        '
        'grid_SlipNoDetail
        '
        Me.grid_SlipNoDetail.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_SlipNoDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_SlipNoDetail.Location = New System.Drawing.Point(0, 48)
        Me.grid_SlipNoDetail.Name = "grid_SlipNoDetail"
        Me.grid_SlipNoDetail.Rows.Count = 2
        Me.grid_SlipNoDetail.Rows.DefaultSize = 20
        Me.grid_SlipNoDetail.Size = New System.Drawing.Size(900, 700)
        Me.grid_SlipNoDetail.StyleInfo = resources.GetString("grid_SlipNoDetail.StyleInfo")
        Me.grid_SlipNoDetail.TabIndex = 10
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Image = Global.Repair_System.My.Resources.Resources.cancel
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(384, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(126, 34)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "불러오기 중지"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = False
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
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Search, Me.ToolStripSeparator1, Me.btn_New_SlipNo, Me.ToolStripSeparator2, Me.btn_Save, Me.Form_CLose})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1264, 25)
        Me.ToolStrip1.TabIndex = 2
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
        'btn_New_SlipNo
        '
        Me.btn_New_SlipNo.Image = Global.Repair_System.My.Resources.Resources.TEST_FOLDER
        Me.btn_New_SlipNo.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_New_SlipNo.Name = "btn_New_SlipNo"
        Me.btn_New_SlipNo.Size = New System.Drawing.Size(75, 22)
        Me.btn_New_SlipNo.Text = "신규등록"
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
        'BTN_File_Load2
        '
        Me.BTN_File_Load2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BTN_File_Load2.Enabled = False
        Me.BTN_File_Load2.Location = New System.Drawing.Point(642, 8)
        Me.BTN_File_Load2.Name = "BTN_File_Load2"
        Me.BTN_File_Load2.Size = New System.Drawing.Size(120, 34)
        Me.BTN_File_Load2.TabIndex = 3
        Me.BTN_File_Load2.Text = "2. Return Type" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "불러오기 (OMS)"
        Me.BTN_File_Load2.UseVisualStyleBackColor = True
        '
        'BTN_File_Load3
        '
        Me.BTN_File_Load3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BTN_File_Load3.Enabled = False
        Me.BTN_File_Load3.Location = New System.Drawing.Point(768, 8)
        Me.BTN_File_Load3.Name = "BTN_File_Load3"
        Me.BTN_File_Load3.Size = New System.Drawing.Size(120, 34)
        Me.BTN_File_Load3.TabIndex = 4
        Me.BTN_File_Load3.Text = "3. Fab, Option" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "불러오기 (우익)"
        Me.BTN_File_Load3.UseVisualStyleBackColor = True
        '
        'frm_Module_In_Add
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_Module_In_Add"
        Me.Text = "Module / Comp 입고 등록"
        Me.cms_Menu.ResumeLayout(False)
        CType(Me.grid_SlipNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grid_SlipNoDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tb_Slip_No As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents dtp_End As DateTimePicker
    Friend WithEvents BTN_File_Load1 As Button
    Friend WithEvents cms_Menu As ContextMenuStrip
    Friend WithEvents btn_RowAdd As ToolStripMenuItem
    Friend WithEvents btn_RowDelete As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents btn_Multi_Row_Insert As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents btn_Delete_Cancel As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents btn_Save2 As ToolStripMenuItem
    Friend WithEvents Label4 As Label
    Friend WithEvents dtp_Start As DateTimePicker
    Friend WithEvents Label1 As Label
    Friend WithEvents grid_SlipNo As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Panel1 As Panel
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Button1 As Button
    Friend WithEvents Form_CLose As ToolStripButton
    Friend WithEvents btn_Save As ToolStripButton
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents btn_New_SlipNo As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents btn_Search As ToolStripButton
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents grid_SlipNoDetail As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents BTN_File_Load3 As Button
    Friend WithEvents BTN_File_Load2 As Button
End Class

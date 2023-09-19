<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Monthly_Production_Data
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Monthly_Production_Data))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.grid_SlipNoDetail = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.tb_StartRow = New System.Windows.Forms.TextBox()
        Me.tb_Product = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.tb_Loss = New System.Windows.Forms.TextBox()
        Me.tb_FabLine = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tb_ReturnType = New System.Windows.Forms.TextBox()
        Me.tb_LotType = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.btn_File_Load = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cms_Menu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.btn_Save2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        CType(Me.grid_SlipNoDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.cms_Menu.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'grid_SlipNoDetail
        '
        Me.grid_SlipNoDetail.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_SlipNoDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_SlipNoDetail.Location = New System.Drawing.Point(0, 135)
        Me.grid_SlipNoDetail.Name = "grid_SlipNoDetail"
        Me.grid_SlipNoDetail.Rows.Count = 2
        Me.grid_SlipNoDetail.Rows.DefaultSize = 20
        Me.grid_SlipNoDetail.Size = New System.Drawing.Size(1264, 638)
        Me.grid_SlipNoDetail.StyleInfo = resources.GetString("grid_SlipNoDetail.StyleInfo")
        Me.grid_SlipNoDetail.TabIndex = 10
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.tb_StartRow)
        Me.GroupBox1.Controls.Add(Me.tb_Product)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.tb_Loss)
        Me.GroupBox1.Controls.Add(Me.tb_FabLine)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.tb_ReturnType)
        Me.GroupBox1.Controls.Add(Me.tb_LotType)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(277, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(872, 55)
        Me.GroupBox1.TabIndex = 28
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Excel 위치설정"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(12, 21)
        Me.Label8.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(92, 21)
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "Product"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_StartRow
        '
        Me.tb_StartRow.BackColor = System.Drawing.SystemColors.Window
        Me.tb_StartRow.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_StartRow.Location = New System.Drawing.Point(814, 21)
        Me.tb_StartRow.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_StartRow.Name = "tb_StartRow"
        Me.tb_StartRow.Size = New System.Drawing.Size(50, 21)
        Me.tb_StartRow.TabIndex = 27
        Me.tb_StartRow.Text = "5"
        Me.tb_StartRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_Product
        '
        Me.tb_Product.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Product.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Product.Location = New System.Drawing.Point(104, 21)
        Me.tb_Product.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Product.Name = "tb_Product"
        Me.tb_Product.Size = New System.Drawing.Size(50, 21)
        Me.tb_Product.TabIndex = 12
        Me.tb_Product.Text = "6"
        Me.tb_Product.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label14.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(722, 21)
        Me.Label14.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(92, 21)
        Me.Label14.TabIndex = 26
        Me.Label14.Text = "시작행"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(154, 21)
        Me.Label7.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(92, 21)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Fab Line"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_Loss
        '
        Me.tb_Loss.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Loss.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Loss.Location = New System.Drawing.Point(672, 21)
        Me.tb_Loss.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Loss.Name = "tb_Loss"
        Me.tb_Loss.Size = New System.Drawing.Size(50, 21)
        Me.tb_Loss.TabIndex = 20
        Me.tb_Loss.Text = "22"
        Me.tb_Loss.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_FabLine
        '
        Me.tb_FabLine.BackColor = System.Drawing.SystemColors.Window
        Me.tb_FabLine.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_FabLine.Location = New System.Drawing.Point(246, 21)
        Me.tb_FabLine.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_FabLine.Name = "tb_FabLine"
        Me.tb_FabLine.Size = New System.Drawing.Size(50, 21)
        Me.tb_FabLine.TabIndex = 14
        Me.tb_FabLine.Text = "7"
        Me.tb_FabLine.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(580, 21)
        Me.Label6.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(92, 21)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Loss(재공)"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(296, 21)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(92, 21)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Lot Type"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_ReturnType
        '
        Me.tb_ReturnType.BackColor = System.Drawing.SystemColors.Window
        Me.tb_ReturnType.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_ReturnType.Location = New System.Drawing.Point(530, 21)
        Me.tb_ReturnType.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_ReturnType.Name = "tb_ReturnType"
        Me.tb_ReturnType.Size = New System.Drawing.Size(50, 21)
        Me.tb_ReturnType.TabIndex = 18
        Me.tb_ReturnType.Text = "9"
        Me.tb_ReturnType.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tb_LotType
        '
        Me.tb_LotType.BackColor = System.Drawing.SystemColors.Window
        Me.tb_LotType.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_LotType.Location = New System.Drawing.Point(388, 21)
        Me.tb_LotType.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_LotType.Name = "tb_LotType"
        Me.tb_LotType.Size = New System.Drawing.Size(50, 21)
        Me.tb_LotType.TabIndex = 16
        Me.tb_LotType.Text = "8"
        Me.tb_LotType.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(438, 21)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(92, 21)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Return Type"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(101, 66)
        Me.DateTimePicker1.Margin = New System.Windows.Forms.Padding(0, 3, 3, 3)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(142, 21)
        Me.DateTimePicker1.TabIndex = 30
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(9, 66)
        Me.Label9.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(92, 21)
        Me.Label9.TabIndex = 29
        Me.Label9.Text = "OMS 일자"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btn_File_Load
        '
        Me.btn_File_Load.Location = New System.Drawing.Point(141, 93)
        Me.btn_File_Load.Name = "btn_File_Load"
        Me.btn_File_Load.Size = New System.Drawing.Size(120, 34)
        Me.btn_File_Load.TabIndex = 4
        Me.btn_File_Load.Text = "OMS 파일" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & " 불러오기"
        Me.btn_File_Load.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label4.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold)
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(0, 25)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(1264, 110)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "OMS 생산월보 저장"
        '
        'cms_Menu
        '
        Me.cms_Menu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Save2})
        Me.cms_Menu.Name = "GRID_MENU"
        Me.cms_Menu.Size = New System.Drawing.Size(99, 26)
        '
        'btn_Save2
        '
        Me.btn_Save2.Image = Global.Repair_System.My.Resources.Resources.save2
        Me.btn_Save2.Name = "btn_Save2"
        Me.btn_Save2.Size = New System.Drawing.Size(98, 22)
        Me.btn_Save2.Text = "저장"
        '
        'Button1
        '
        Me.Button1.Image = Global.Repair_System.My.Resources.Resources.cancel
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(9, 93)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(126, 34)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "불러오기 중지"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = False
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Form_CLose})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1264, 25)
        Me.ToolStrip1.TabIndex = 31
        Me.ToolStrip1.Text = "ToolStrip1"
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
        'frm_Monthly_Production_Data
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.grid_SlipNoDetail)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btn_File_Load)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_Monthly_Production_Data"
        Me.Text = "OMS 생산월보 저장"
        CType(Me.grid_SlipNoDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.cms_Menu.ResumeLayout(False)
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents grid_SlipNoDetail As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Button1 As Button
    Friend WithEvents btn_File_Load As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents cms_Menu As ContextMenuStrip
    Friend WithEvents tb_Loss As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents tb_ReturnType As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents tb_LotType As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents tb_FabLine As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents tb_Product As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents tb_StartRow As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents Label9 As Label
    Friend WithEvents btn_Save2 As ToolStripMenuItem
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents Form_CLose As ToolStripButton
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_oms_file_data
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_oms_file_data))
        Me.btn_Menu_Save = New System.Windows.Forms.ToolStripButton()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.btn_Search = New System.Windows.Forms.ToolStripButton()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.TB_File_Type = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TB_File_Desc = New System.Windows.Forms.TextBox()
        Me.TB_Frequency = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TB_IF_Name = New System.Windows.Forms.TextBox()
        Me.grid_Data_List = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.btn_SAVE = New System.Windows.Forms.ToolStripMenuItem()
        Me.COL_VIEW = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Row_Delete = New System.Windows.Forms.ToolStripMenuItem()
        Me.btn_Row_Add = New System.Windows.Forms.ToolStripMenuItem()
        Me.GRID_Menu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStrip1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grid_Data_List, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GRID_Menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_Menu_Save
        '
        Me.btn_Menu_Save.Image = Global.Repair_System.My.Resources.Resources.save4
        Me.btn_Menu_Save.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_Menu_Save.Name = "btn_Menu_Save"
        Me.btn_Menu_Save.Size = New System.Drawing.Size(51, 22)
        Me.btn_Menu_Save.Text = "저장"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Search, Me.btn_Menu_Save})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1264, 25)
        Me.ToolStrip1.TabIndex = 54
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
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.SplitContainer1.Panel1.Controls.Add(Me.ToolStrip1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label31)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TB_File_Type)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TB_File_Desc)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TB_Frequency)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TB_IF_Name)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.White
        Me.SplitContainer1.Panel2.Controls.Add(Me.grid_Data_List)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 773)
        Me.SplitContainer1.SplitterDistance = 90
        Me.SplitContainer1.TabIndex = 56
        '
        'Label31
        '
        Me.Label31.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label31.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label31.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label31.ForeColor = System.Drawing.Color.White
        Me.Label31.Location = New System.Drawing.Point(9, 35)
        Me.Label31.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(170, 21)
        Me.Label31.TabIndex = 46
        Me.Label31.Text = "File Type"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_File_Type
        '
        Me.TB_File_Type.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.TB_File_Type.Font = New System.Drawing.Font("굴림", 9.75!)
        Me.TB_File_Type.FormattingEnabled = True
        Me.TB_File_Type.Items.AddRange(New Object() {"CMWIPINF", "CMHALFIN"})
        Me.TB_File_Type.Location = New System.Drawing.Point(179, 35)
        Me.TB_File_Type.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_File_Type.Name = "TB_File_Type"
        Me.TB_File_Type.Size = New System.Drawing.Size(297, 21)
        Me.TB_File_Type.TabIndex = 47
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(476, 35)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(170, 21)
        Me.Label3.TabIndex = 48
        Me.Label3.Text = "File Description"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_File_Desc
        '
        Me.TB_File_Desc.BackColor = System.Drawing.SystemColors.Window
        Me.TB_File_Desc.Enabled = False
        Me.TB_File_Desc.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_File_Desc.Location = New System.Drawing.Point(646, 35)
        Me.TB_File_Desc.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_File_Desc.Name = "TB_File_Desc"
        Me.TB_File_Desc.Size = New System.Drawing.Size(297, 21)
        Me.TB_File_Desc.TabIndex = 49
        '
        'TB_Frequency
        '
        Me.TB_Frequency.BackColor = System.Drawing.SystemColors.Window
        Me.TB_Frequency.Enabled = False
        Me.TB_Frequency.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_Frequency.Location = New System.Drawing.Point(646, 58)
        Me.TB_Frequency.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_Frequency.Name = "TB_Frequency"
        Me.TB_Frequency.Size = New System.Drawing.Size(297, 21)
        Me.TB_Frequency.TabIndex = 53
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(9, 58)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(170, 21)
        Me.Label1.TabIndex = 50
        Me.Label1.Text = "I/F Name"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(476, 58)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(170, 21)
        Me.Label2.TabIndex = 52
        Me.Label2.Text = "Frequency"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TB_IF_Name
        '
        Me.TB_IF_Name.BackColor = System.Drawing.SystemColors.Window
        Me.TB_IF_Name.Enabled = False
        Me.TB_IF_Name.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TB_IF_Name.Location = New System.Drawing.Point(179, 58)
        Me.TB_IF_Name.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_IF_Name.Name = "TB_IF_Name"
        Me.TB_IF_Name.Size = New System.Drawing.Size(297, 21)
        Me.TB_IF_Name.TabIndex = 51
        '
        'grid_Data_List
        '
        Me.grid_Data_List.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.grid_Data_List.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grid_Data_List.Location = New System.Drawing.Point(0, 0)
        Me.grid_Data_List.Name = "grid_Data_List"
        Me.grid_Data_List.Rows.Count = 2
        Me.grid_Data_List.Rows.DefaultSize = 20
        Me.grid_Data_List.Size = New System.Drawing.Size(1264, 679)
        Me.grid_Data_List.StyleInfo = resources.GetString("grid_Data_List.StyleInfo")
        Me.grid_Data_List.TabIndex = 54
        '
        'btn_SAVE
        '
        Me.btn_SAVE.Name = "btn_SAVE"
        Me.btn_SAVE.Size = New System.Drawing.Size(114, 22)
        Me.btn_SAVE.Text = "저장"
        '
        'COL_VIEW
        '
        Me.COL_VIEW.Name = "COL_VIEW"
        Me.COL_VIEW.Size = New System.Drawing.Size(111, 6)
        '
        'btn_Row_Delete
        '
        Me.btn_Row_Delete.Name = "btn_Row_Delete"
        Me.btn_Row_Delete.Size = New System.Drawing.Size(114, 22)
        Me.btn_Row_Delete.Text = "행 삭제"
        '
        'btn_Row_Add
        '
        Me.btn_Row_Add.Name = "btn_Row_Add"
        Me.btn_Row_Add.Size = New System.Drawing.Size(114, 22)
        Me.btn_Row_Add.Text = "행 추가"
        '
        'GRID_Menu
        '
        Me.GRID_Menu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Row_Add, Me.btn_Row_Delete, Me.COL_VIEW, Me.btn_SAVE})
        Me.GRID_Menu.Name = "GRID_Menu"
        Me.GRID_Menu.Size = New System.Drawing.Size(115, 76)
        '
        'frm_oms_file_data
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "frm_oms_file_data"
        Me.Text = "OMS 파일 기초정보"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grid_Data_List, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GRID_Menu.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btn_Menu_Save As ToolStripButton
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents btn_Search As ToolStripButton
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Label31 As Label
    Friend WithEvents TB_File_Type As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents TB_File_Desc As TextBox
    Friend WithEvents TB_Frequency As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents TB_IF_Name As TextBox
    Friend WithEvents grid_Data_List As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents btn_SAVE As ToolStripMenuItem
    Friend WithEvents COL_VIEW As ToolStripSeparator
    Friend WithEvents btn_Row_Delete As ToolStripMenuItem
    Friend WithEvents btn_Row_Add As ToolStripMenuItem
    Friend WithEvents GRID_Menu As ContextMenuStrip
End Class

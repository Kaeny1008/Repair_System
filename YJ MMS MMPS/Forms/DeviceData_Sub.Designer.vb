<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DeviceData_Sub
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DeviceData_Sub))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Tb_Customer = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Tb_mainParts = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Tb_DDCode = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Tb_feederNo = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Tb_machineNo = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Grid_DeviceData = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.CMS_gridMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.BTN_newLine = New System.Windows.Forms.ToolStripMenuItem()
        Me.BTN_deleteLine = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BTN_Save = New System.Windows.Forms.ToolStripMenuItem()
        Me.TB_mainParts_Maker = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.Grid_DeviceData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CMS_gridMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.TB_mainParts_Maker)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Tb_Customer)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Tb_mainParts)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Tb_DDCode)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Tb_feederNo)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Tb_machineNo)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(776, 115)
        Me.Panel1.TabIndex = 1
        '
        'Tb_Customer
        '
        Me.Tb_Customer.Enabled = False
        Me.Tb_Customer.Location = New System.Drawing.Point(178, 33)
        Me.Tb_Customer.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_Customer.Name = "Tb_Customer"
        Me.Tb_Customer.Size = New System.Drawing.Size(510, 21)
        Me.Tb_Customer.TabIndex = 12
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.SlateGray
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(68, 33)
        Me.Label6.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(110, 21)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "고객사"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Tb_mainParts
        '
        Me.Tb_mainParts.Enabled = False
        Me.Tb_mainParts.Location = New System.Drawing.Point(178, 79)
        Me.Tb_mainParts.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_mainParts.Name = "Tb_mainParts"
        Me.Tb_mainParts.Size = New System.Drawing.Size(180, 21)
        Me.Tb_mainParts.TabIndex = 10
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.SlateGray
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(68, 79)
        Me.Label5.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(110, 21)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "대표 Part No."
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Tb_DDCode
        '
        Me.Tb_DDCode.Enabled = False
        Me.Tb_DDCode.Location = New System.Drawing.Point(538, 56)
        Me.Tb_DDCode.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_DDCode.Name = "Tb_DDCode"
        Me.Tb_DDCode.Size = New System.Drawing.Size(150, 21)
        Me.Tb_DDCode.TabIndex = 8
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.SlateGray
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(428, 56)
        Me.Label4.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(110, 21)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "DD Code"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Tb_feederNo
        '
        Me.Tb_feederNo.Enabled = False
        Me.Tb_feederNo.Location = New System.Drawing.Point(358, 56)
        Me.Tb_feederNo.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_feederNo.Name = "Tb_feederNo"
        Me.Tb_feederNo.Size = New System.Drawing.Size(70, 21)
        Me.Tb_feederNo.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.SlateGray
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(248, 56)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(110, 21)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Feeder No."
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Tb_machineNo
        '
        Me.Tb_machineNo.Enabled = False
        Me.Tb_machineNo.Location = New System.Drawing.Point(178, 56)
        Me.Tb_machineNo.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Tb_machineNo.Name = "Tb_machineNo"
        Me.Tb_machineNo.Size = New System.Drawing.Size(70, 21)
        Me.Tb_machineNo.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.SlateGray
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(68, 56)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(110, 21)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Machine No."
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(3, 3)
        Me.Label1.Margin = New System.Windows.Forms.Padding(3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(766, 26)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "[ 등록정보 ]"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Grid_DeviceData
        '
        Me.Grid_DeviceData.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.Grid_DeviceData.Location = New System.Drawing.Point(12, 133)
        Me.Grid_DeviceData.Name = "Grid_DeviceData"
        Me.Grid_DeviceData.Rows.Count = 2
        Me.Grid_DeviceData.Rows.DefaultSize = 20
        Me.Grid_DeviceData.Size = New System.Drawing.Size(776, 335)
        Me.Grid_DeviceData.StyleInfo = resources.GetString("Grid_DeviceData.StyleInfo")
        Me.Grid_DeviceData.TabIndex = 5
        Me.Grid_DeviceData.VisualStyle = C1.Win.C1FlexGrid.VisualStyle.Office2010Blue
        '
        'CMS_gridMenu
        '
        Me.CMS_gridMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_newLine, Me.BTN_deleteLine, Me.ToolStripSeparator1, Me.BTN_Save})
        Me.CMS_gridMenu.Name = "CMS_gridMenu"
        Me.CMS_gridMenu.Size = New System.Drawing.Size(123, 76)
        '
        'BTN_newLine
        '
        Me.BTN_newLine.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.plus_blue
        Me.BTN_newLine.Name = "BTN_newLine"
        Me.BTN_newLine.Size = New System.Drawing.Size(122, 22)
        Me.BTN_newLine.Text = "신규등록"
        '
        'BTN_deleteLine
        '
        Me.BTN_deleteLine.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.minus_blue
        Me.BTN_deleteLine.Name = "BTN_deleteLine"
        Me.BTN_deleteLine.Size = New System.Drawing.Size(122, 22)
        Me.BTN_deleteLine.Text = "삭제"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(119, 6)
        '
        'BTN_Save
        '
        Me.BTN_Save.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.save_5
        Me.BTN_Save.Name = "BTN_Save"
        Me.BTN_Save.Size = New System.Drawing.Size(122, 22)
        Me.BTN_Save.Text = "저장"
        '
        'TB_mainParts_Maker
        '
        Me.TB_mainParts_Maker.Enabled = False
        Me.TB_mainParts_Maker.Location = New System.Drawing.Point(508, 79)
        Me.TB_mainParts_Maker.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_mainParts_Maker.Name = "TB_mainParts_Maker"
        Me.TB_mainParts_Maker.Size = New System.Drawing.Size(180, 21)
        Me.TB_mainParts_Maker.TabIndex = 14
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.SlateGray
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(358, 79)
        Me.Label7.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(150, 21)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "대표 Part No. Maker"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'DeviceData_Sub
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.LightSteelBlue
        Me.ClientSize = New System.Drawing.Size(800, 482)
        Me.Controls.Add(Me.Grid_DeviceData)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DeviceData_Sub"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "대치자재 관리"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.Grid_DeviceData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CMS_gridMenu.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents Tb_mainParts As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Tb_DDCode As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Tb_feederNo As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Tb_machineNo As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Grid_DeviceData As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents CMS_gridMenu As ContextMenuStrip
    Friend WithEvents BTN_newLine As ToolStripMenuItem
    Friend WithEvents BTN_deleteLine As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents BTN_Save As ToolStripMenuItem
    Friend WithEvents Tb_Customer As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents TB_mainParts_Maker As TextBox
    Friend WithEvents Label7 As Label
End Class

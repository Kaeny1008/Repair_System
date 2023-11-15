<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Mainform
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Mainform))
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.BTN_Home = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.Btn_CodeManager = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripLabel2 = New System.Windows.Forms.ToolStripLabel()
        Me.Btn_ModelManager = New System.Windows.Forms.ToolStripButton()
        Me.BTN_LabelPrint = New System.Windows.Forms.ToolStripButton()
        Me.BTN_HistoryBLU = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.Btn_DDW = New System.Windows.Forms.ToolStripButton()
        Me.Btn_History = New System.Windows.Forms.ToolStripButton()
        Me.BTN_Form_Close = New System.Windows.Forms.ToolStripButton()
        Me.ToolStrip = New System.Windows.Forms.ToolStrip()
        Me.StatusStrip.SuspendLayout()
        Me.ToolStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 725)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Padding = New System.Windows.Forms.Padding(1, 0, 16, 0)
        Me.StatusStrip.Size = New System.Drawing.Size(1209, 22)
        Me.StatusStrip.TabIndex = 7
        Me.StatusStrip.Text = "StatusStrip"
        '
        'ToolStripStatusLabel
        '
        Me.ToolStripStatusLabel.Name = "ToolStripStatusLabel"
        Me.ToolStripStatusLabel.Size = New System.Drawing.Size(31, 17)
        Me.ToolStripStatusLabel.Text = "상태"
        '
        'MenuStrip
        '
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Padding = New System.Windows.Forms.Padding(7, 2, 0, 2)
        Me.MenuStrip.Size = New System.Drawing.Size(1209, 24)
        Me.MenuStrip.TabIndex = 5
        Me.MenuStrip.Text = "MenuStrip"
        Me.MenuStrip.Visible = False
        '
        'BTN_Home
        '
        Me.BTN_Home.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_Home.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.home
        Me.BTN_Home.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_Home.Name = "BTN_Home"
        Me.BTN_Home.Size = New System.Drawing.Size(80, 28)
        Me.BTN_Home.Text = "Home"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 31)
        '
        'Btn_CodeManager
        '
        Me.Btn_CodeManager.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.Btn_CodeManager.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.DeviceData2
        Me.Btn_CodeManager.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.Btn_CodeManager.Name = "Btn_CodeManager"
        Me.Btn_CodeManager.Size = New System.Drawing.Size(108, 28)
        Me.Btn_CodeManager.Text = "Code 관리"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(6, 31)
        '
        'ToolStripLabel2
        '
        Me.ToolStripLabel2.AutoSize = False
        Me.ToolStripLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripLabel2.Font = New System.Drawing.Font("맑은 고딕", 11.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.ToolStripLabel2.ForeColor = System.Drawing.SystemColors.ButtonShadow
        Me.ToolStripLabel2.Name = "ToolStripLabel2"
        Me.ToolStripLabel2.Size = New System.Drawing.Size(100, 28)
        Me.ToolStripLabel2.Text = "[ BLU 관련 ]"
        '
        'Btn_ModelManager
        '
        Me.Btn_ModelManager.Enabled = False
        Me.Btn_ModelManager.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.Btn_ModelManager.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.folder
        Me.Btn_ModelManager.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.Btn_ModelManager.Name = "Btn_ModelManager"
        Me.Btn_ModelManager.Size = New System.Drawing.Size(149, 28)
        Me.Btn_ModelManager.Text = "모델 등록 / 관리"
        '
        'BTN_LabelPrint
        '
        Me.BTN_LabelPrint.Enabled = False
        Me.BTN_LabelPrint.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.BTN_LabelPrint.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.printer_blue
        Me.BTN_LabelPrint.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_LabelPrint.Name = "BTN_LabelPrint"
        Me.BTN_LabelPrint.Size = New System.Drawing.Size(109, 28)
        Me.BTN_LabelPrint.Text = "Label 발행"
        '
        'BTN_HistoryBLU
        '
        Me.BTN_HistoryBLU.Enabled = False
        Me.BTN_HistoryBLU.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.BTN_HistoryBLU.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.vmware
        Me.BTN_HistoryBLU.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_HistoryBLU.Name = "BTN_HistoryBLU"
        Me.BTN_HistoryBLU.Size = New System.Drawing.Size(97, 28)
        Me.BTN_HistoryBLU.Text = "이력보기"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 31)
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.AutoSize = False
        Me.ToolStripLabel1.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripLabel1.Font = New System.Drawing.Font("맑은 고딕", 11.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.ToolStripLabel1.ForeColor = System.Drawing.SystemColors.ButtonShadow
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(100, 28)
        Me.ToolStripLabel1.Text = "[ BLU 외 ]"
        '
        'Btn_DDW
        '
        Me.Btn_DDW.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.Btn_DDW.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.DeviceData1
        Me.Btn_DDW.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.Btn_DDW.Name = "Btn_DDW"
        Me.Btn_DDW.Size = New System.Drawing.Size(151, 28)
        Me.Btn_DDW.Text = "DeviceData 작성"
        '
        'Btn_History
        '
        Me.Btn_History.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.Btn_History.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.vmware
        Me.Btn_History.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.Btn_History.Name = "Btn_History"
        Me.Btn_History.Size = New System.Drawing.Size(97, 28)
        Me.Btn_History.Text = "이력보기"
        '
        'BTN_Form_Close
        '
        Me.BTN_Form_Close.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.BTN_Form_Close.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BTN_Form_Close.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold)
        Me.BTN_Form_Close.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.Close
        Me.BTN_Form_Close.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BTN_Form_Close.Name = "BTN_Form_Close"
        Me.BTN_Form_Close.Size = New System.Drawing.Size(28, 28)
        Me.BTN_Form_Close.Text = "닫기"
        '
        'ToolStrip
        '
        Me.ToolStrip.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BTN_Home, Me.ToolStripSeparator2, Me.Btn_CodeManager, Me.ToolStripSeparator5, Me.ToolStripLabel2, Me.Btn_ModelManager, Me.BTN_LabelPrint, Me.BTN_HistoryBLU, Me.ToolStripSeparator3, Me.ToolStripLabel1, Me.Btn_DDW, Me.Btn_History, Me.BTN_Form_Close})
        Me.ToolStrip.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip.Name = "ToolStrip"
        Me.ToolStrip.Size = New System.Drawing.Size(1209, 31)
        Me.ToolStrip.TabIndex = 6
        Me.ToolStrip.Text = "ToolStrip"
        '
        'Mainform
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1209, 747)
        Me.Controls.Add(Me.ToolStrip)
        Me.Controls.Add(Me.MenuStrip)
        Me.Controls.Add(Me.StatusStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "Mainform"
        Me.Text = "Yujin Manufacturing Management System : 오삽방지 시스템"
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ToolStrip.ResumeLayout(False)
        Me.ToolStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents MenuStrip As MenuStrip
    Friend WithEvents BTN_Home As ToolStripButton
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents Btn_CodeManager As ToolStripButton
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents ToolStripLabel2 As ToolStripLabel
    Friend WithEvents Btn_ModelManager As ToolStripButton
    Friend WithEvents BTN_LabelPrint As ToolStripButton
    Friend WithEvents BTN_HistoryBLU As ToolStripButton
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents ToolStripLabel1 As ToolStripLabel
    Friend WithEvents Btn_DDW As ToolStripButton
    Friend WithEvents Btn_History As ToolStripButton
    Friend WithEvents BTN_Form_Close As ToolStripButton
    Friend WithEvents ToolStrip As ToolStrip
End Class

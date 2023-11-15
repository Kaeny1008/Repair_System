<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Subform
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
        Me.BTN_uploadPG = New System.Windows.Forms.Button()
        Me.BTN_connection = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.BTN_HistoryBLU = New System.Windows.Forms.Button()
        Me.BTN_LabelPrint = New System.Windows.Forms.Button()
        Me.Btn_ModelManager = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Btn_DDW = New System.Windows.Forms.Button()
        Me.Btn_History = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Btn_CodeManager = New System.Windows.Forms.Button()
        Me.Btn_Exit = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BTN_uploadPG
        '
        Me.BTN_uploadPG.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BTN_uploadPG.Location = New System.Drawing.Point(93, 741)
        Me.BTN_uploadPG.Name = "BTN_uploadPG"
        Me.BTN_uploadPG.Size = New System.Drawing.Size(75, 23)
        Me.BTN_uploadPG.TabIndex = 30
        Me.BTN_uploadPG.Text = "PG Upload"
        Me.BTN_uploadPG.UseVisualStyleBackColor = True
        '
        'BTN_connection
        '
        Me.BTN_connection.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BTN_connection.Location = New System.Drawing.Point(12, 741)
        Me.BTN_connection.Name = "BTN_connection"
        Me.BTN_connection.Size = New System.Drawing.Size(75, 23)
        Me.BTN_connection.TabIndex = 29
        Me.BTN_connection.Text = "접속정보"
        Me.BTN_connection.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.BTN_HistoryBLU)
        Me.GroupBox1.Controls.Add(Me.BTN_LabelPrint)
        Me.GroupBox1.Controls.Add(Me.Btn_ModelManager)
        Me.GroupBox1.Enabled = False
        Me.GroupBox1.Font = New System.Drawing.Font("굴림", 24.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1240, 231)
        Me.GroupBox1.TabIndex = 32
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "[ BLU 관련 ]"
        '
        'BTN_HistoryBLU
        '
        Me.BTN_HistoryBLU.BackColor = System.Drawing.Color.SandyBrown
        Me.BTN_HistoryBLU.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_HistoryBLU.ForeColor = System.Drawing.Color.White
        Me.BTN_HistoryBLU.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.vmware
        Me.BTN_HistoryBLU.Location = New System.Drawing.Point(526, 43)
        Me.BTN_HistoryBLU.Name = "BTN_HistoryBLU"
        Me.BTN_HistoryBLU.Size = New System.Drawing.Size(254, 178)
        Me.BTN_HistoryBLU.TabIndex = 33
        Me.BTN_HistoryBLU.Text = "이력보기"
        Me.BTN_HistoryBLU.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BTN_HistoryBLU.UseVisualStyleBackColor = False
        '
        'BTN_LabelPrint
        '
        Me.BTN_LabelPrint.BackColor = System.Drawing.Color.Peru
        Me.BTN_LabelPrint.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_LabelPrint.ForeColor = System.Drawing.Color.White
        Me.BTN_LabelPrint.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.printer_blue
        Me.BTN_LabelPrint.Location = New System.Drawing.Point(266, 43)
        Me.BTN_LabelPrint.Name = "BTN_LabelPrint"
        Me.BTN_LabelPrint.Size = New System.Drawing.Size(254, 178)
        Me.BTN_LabelPrint.TabIndex = 32
        Me.BTN_LabelPrint.Text = "Label 발행"
        Me.BTN_LabelPrint.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BTN_LabelPrint.UseVisualStyleBackColor = False
        '
        'Btn_ModelManager
        '
        Me.Btn_ModelManager.BackColor = System.Drawing.Color.BurlyWood
        Me.Btn_ModelManager.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_ModelManager.ForeColor = System.Drawing.Color.White
        Me.Btn_ModelManager.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.folder
        Me.Btn_ModelManager.Location = New System.Drawing.Point(6, 43)
        Me.Btn_ModelManager.Name = "Btn_ModelManager"
        Me.Btn_ModelManager.Size = New System.Drawing.Size(254, 178)
        Me.Btn_ModelManager.TabIndex = 31
        Me.Btn_ModelManager.Text = "모델 등록 / 관리"
        Me.Btn_ModelManager.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Btn_ModelManager.UseVisualStyleBackColor = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.Btn_DDW)
        Me.GroupBox2.Controls.Add(Me.Btn_History)
        Me.GroupBox2.Font = New System.Drawing.Font("굴림", 24.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(12, 308)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1240, 231)
        Me.GroupBox2.TabIndex = 33
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "[ BLU 외 ]"
        '
        'Btn_DDW
        '
        Me.Btn_DDW.BackColor = System.Drawing.Color.SaddleBrown
        Me.Btn_DDW.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_DDW.ForeColor = System.Drawing.Color.White
        Me.Btn_DDW.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.DeviceData1
        Me.Btn_DDW.Location = New System.Drawing.Point(6, 43)
        Me.Btn_DDW.Name = "Btn_DDW"
        Me.Btn_DDW.Size = New System.Drawing.Size(254, 178)
        Me.Btn_DDW.TabIndex = 16
        Me.Btn_DDW.Text = "DeviceData 작성"
        Me.Btn_DDW.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Btn_DDW.UseVisualStyleBackColor = False
        '
        'Btn_History
        '
        Me.Btn_History.BackColor = System.Drawing.Color.SteelBlue
        Me.Btn_History.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_History.ForeColor = System.Drawing.Color.White
        Me.Btn_History.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.vmware
        Me.Btn_History.Location = New System.Drawing.Point(266, 43)
        Me.Btn_History.Name = "Btn_History"
        Me.Btn_History.Size = New System.Drawing.Size(254, 178)
        Me.Btn_History.TabIndex = 21
        Me.Btn_History.Text = "이력보기"
        Me.Btn_History.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Btn_History.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.YUJIN_Logo
        Me.PictureBox1.Location = New System.Drawing.Point(589, 566)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(666, 198)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox1.TabIndex = 23
        Me.PictureBox1.TabStop = False
        '
        'Btn_CodeManager
        '
        Me.Btn_CodeManager.BackColor = System.Drawing.Color.DarkGoldenrod
        Me.Btn_CodeManager.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_CodeManager.ForeColor = System.Drawing.Color.White
        Me.Btn_CodeManager.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.DeviceData2
        Me.Btn_CodeManager.Location = New System.Drawing.Point(589, 545)
        Me.Btn_CodeManager.Name = "Btn_CodeManager"
        Me.Btn_CodeManager.Size = New System.Drawing.Size(254, 178)
        Me.Btn_CodeManager.TabIndex = 24
        Me.Btn_CodeManager.Text = "Code 관리"
        Me.Btn_CodeManager.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Btn_CodeManager.UseVisualStyleBackColor = False
        Me.Btn_CodeManager.Visible = False
        '
        'Btn_Exit
        '
        Me.Btn_Exit.BackColor = System.Drawing.Color.DarkMagenta
        Me.Btn_Exit.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Btn_Exit.ForeColor = System.Drawing.Color.White
        Me.Btn_Exit.Image = Global.YJ_MMS_MMPS.My.Resources.Resources.Close_Box_Red
        Me.Btn_Exit.Location = New System.Drawing.Point(698, 545)
        Me.Btn_Exit.Name = "Btn_Exit"
        Me.Btn_Exit.Size = New System.Drawing.Size(254, 178)
        Me.Btn_Exit.TabIndex = 22
        Me.Btn_Exit.Text = "종료"
        Me.Btn_Exit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Btn_Exit.UseVisualStyleBackColor = False
        Me.Btn_Exit.Visible = False
        '
        'Subform
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BTN_uploadPG)
        Me.Controls.Add(Me.BTN_connection)
        Me.Controls.Add(Me.Btn_CodeManager)
        Me.Controls.Add(Me.Btn_Exit)
        Me.Name = "Subform"
        Me.Text = "Home"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Btn_DDW As Button
    Friend WithEvents Btn_History As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents BTN_uploadPG As Button
    Friend WithEvents BTN_connection As Button
    Friend WithEvents Btn_ModelManager As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Btn_CodeManager As Button
    Friend WithEvents Btn_Exit As Button
    Friend WithEvents BTN_LabelPrint As Button
    Friend WithEvents BTN_HistoryBLU As Button
End Class

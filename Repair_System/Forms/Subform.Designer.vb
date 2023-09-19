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
        Me.BTN_connection = New System.Windows.Forms.Button()
        Me.BTN_uploadPG = New System.Windows.Forms.Button()
        Me.BTN_Cal_Object = New System.Windows.Forms.Button()
        Me.BTN_Asset_Inspection = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.BTN_Close = New System.Windows.Forms.Button()
        Me.BTN_Asset_History = New System.Windows.Forms.Button()
        Me.BTN_Asset_Manager2 = New System.Windows.Forms.Button()
        Me.BTN_Asset_Manager = New System.Windows.Forms.Button()
        Me.BTN_Asset_ADD = New System.Windows.Forms.Button()
        Me.BTN_Code_Manager = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BTN_connection
        '
        Me.BTN_connection.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BTN_connection.Location = New System.Drawing.Point(12, 738)
        Me.BTN_connection.Name = "BTN_connection"
        Me.BTN_connection.Size = New System.Drawing.Size(75, 23)
        Me.BTN_connection.TabIndex = 27
        Me.BTN_connection.Text = "접속정보"
        Me.BTN_connection.UseVisualStyleBackColor = True
        '
        'BTN_uploadPG
        '
        Me.BTN_uploadPG.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BTN_uploadPG.Location = New System.Drawing.Point(93, 738)
        Me.BTN_uploadPG.Name = "BTN_uploadPG"
        Me.BTN_uploadPG.Size = New System.Drawing.Size(75, 23)
        Me.BTN_uploadPG.TabIndex = 28
        Me.BTN_uploadPG.Text = "PG Upload"
        Me.BTN_uploadPG.UseVisualStyleBackColor = True
        '
        'BTN_Cal_Object
        '
        Me.BTN_Cal_Object.BackColor = System.Drawing.Color.DarkBlue
        Me.BTN_Cal_Object.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_Cal_Object.ForeColor = System.Drawing.Color.White
        Me.BTN_Cal_Object.Image = Global.YJ_MMS_Asset.My.Resources.Resources.Process_Info_128
        Me.BTN_Cal_Object.Location = New System.Drawing.Point(532, 196)
        Me.BTN_Cal_Object.Name = "BTN_Cal_Object"
        Me.BTN_Cal_Object.Size = New System.Drawing.Size(254, 178)
        Me.BTN_Cal_Object.TabIndex = 29
        Me.BTN_Cal_Object.Text = "검교정 관리"
        Me.BTN_Cal_Object.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTN_Cal_Object.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BTN_Cal_Object.UseVisualStyleBackColor = False
        '
        'BTN_Asset_Inspection
        '
        Me.BTN_Asset_Inspection.BackColor = System.Drawing.Color.CadetBlue
        Me.BTN_Asset_Inspection.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_Asset_Inspection.ForeColor = System.Drawing.Color.White
        Me.BTN_Asset_Inspection.Image = Global.YJ_MMS_Asset.My.Resources.Resources.ordering_128
        Me.BTN_Asset_Inspection.Location = New System.Drawing.Point(272, 196)
        Me.BTN_Asset_Inspection.Name = "BTN_Asset_Inspection"
        Me.BTN_Asset_Inspection.Size = New System.Drawing.Size(254, 178)
        Me.BTN_Asset_Inspection.TabIndex = 26
        Me.BTN_Asset_Inspection.Text = "재물조사"
        Me.BTN_Asset_Inspection.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTN_Asset_Inspection.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BTN_Asset_Inspection.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Image = Global.YJ_MMS_Asset.My.Resources.Resources.YUJIN_Logo
        Me.PictureBox1.Location = New System.Drawing.Point(589, 566)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(666, 198)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox1.TabIndex = 25
        Me.PictureBox1.TabStop = False
        '
        'BTN_Close
        '
        Me.BTN_Close.BackColor = System.Drawing.Color.DarkMagenta
        Me.BTN_Close.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_Close.ForeColor = System.Drawing.Color.White
        Me.BTN_Close.Image = Global.YJ_MMS_Asset.My.Resources.Resources.Close_Box_Red
        Me.BTN_Close.Location = New System.Drawing.Point(792, 196)
        Me.BTN_Close.Name = "BTN_Close"
        Me.BTN_Close.Size = New System.Drawing.Size(254, 178)
        Me.BTN_Close.TabIndex = 24
        Me.BTN_Close.Text = "종료"
        Me.BTN_Close.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTN_Close.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BTN_Close.UseVisualStyleBackColor = False
        '
        'BTN_Asset_History
        '
        Me.BTN_Asset_History.BackColor = System.Drawing.Color.SteelBlue
        Me.BTN_Asset_History.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_Asset_History.ForeColor = System.Drawing.Color.White
        Me.BTN_Asset_History.Image = Global.YJ_MMS_Asset.My.Resources.Resources.Star_Full_128
        Me.BTN_Asset_History.Location = New System.Drawing.Point(12, 196)
        Me.BTN_Asset_History.Name = "BTN_Asset_History"
        Me.BTN_Asset_History.Size = New System.Drawing.Size(254, 178)
        Me.BTN_Asset_History.TabIndex = 23
        Me.BTN_Asset_History.Text = "자산현황"
        Me.BTN_Asset_History.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTN_Asset_History.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BTN_Asset_History.UseVisualStyleBackColor = False
        '
        'BTN_Asset_Manager2
        '
        Me.BTN_Asset_Manager2.BackColor = System.Drawing.Color.DarkOliveGreen
        Me.BTN_Asset_Manager2.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_Asset_Manager2.ForeColor = System.Drawing.Color.White
        Me.BTN_Asset_Manager2.Image = Global.YJ_MMS_Asset.My.Resources.Resources.Globe_Download_128
        Me.BTN_Asset_Manager2.Location = New System.Drawing.Point(792, 12)
        Me.BTN_Asset_Manager2.Name = "BTN_Asset_Manager2"
        Me.BTN_Asset_Manager2.Size = New System.Drawing.Size(254, 178)
        Me.BTN_Asset_Manager2.TabIndex = 22
        Me.BTN_Asset_Manager2.Text = "자산관리(일괄)"
        Me.BTN_Asset_Manager2.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTN_Asset_Manager2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BTN_Asset_Manager2.UseVisualStyleBackColor = False
        '
        'BTN_Asset_Manager
        '
        Me.BTN_Asset_Manager.BackColor = System.Drawing.Color.Green
        Me.BTN_Asset_Manager.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_Asset_Manager.ForeColor = System.Drawing.Color.White
        Me.BTN_Asset_Manager.Image = Global.YJ_MMS_Asset.My.Resources.Resources.web_management
        Me.BTN_Asset_Manager.Location = New System.Drawing.Point(532, 12)
        Me.BTN_Asset_Manager.Name = "BTN_Asset_Manager"
        Me.BTN_Asset_Manager.Size = New System.Drawing.Size(254, 178)
        Me.BTN_Asset_Manager.TabIndex = 22
        Me.BTN_Asset_Manager.Text = "자산관리(개별)"
        Me.BTN_Asset_Manager.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTN_Asset_Manager.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BTN_Asset_Manager.UseVisualStyleBackColor = False
        '
        'BTN_Asset_ADD
        '
        Me.BTN_Asset_ADD.BackColor = System.Drawing.Color.DarkGoldenrod
        Me.BTN_Asset_ADD.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_Asset_ADD.ForeColor = System.Drawing.Color.White
        Me.BTN_Asset_ADD.Image = Global.YJ_MMS_Asset.My.Resources.Resources.Open_Folder_Add
        Me.BTN_Asset_ADD.Location = New System.Drawing.Point(272, 12)
        Me.BTN_Asset_ADD.Name = "BTN_Asset_ADD"
        Me.BTN_Asset_ADD.Size = New System.Drawing.Size(254, 178)
        Me.BTN_Asset_ADD.TabIndex = 21
        Me.BTN_Asset_ADD.Text = "자산등록"
        Me.BTN_Asset_ADD.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTN_Asset_ADD.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BTN_Asset_ADD.UseVisualStyleBackColor = False
        '
        'BTN_Code_Manager
        '
        Me.BTN_Code_Manager.BackColor = System.Drawing.Color.SaddleBrown
        Me.BTN_Code_Manager.Font = New System.Drawing.Font("굴림", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BTN_Code_Manager.ForeColor = System.Drawing.Color.White
        Me.BTN_Code_Manager.Image = Global.YJ_MMS_Asset.My.Resources.Resources.product_sales_report_128
        Me.BTN_Code_Manager.Location = New System.Drawing.Point(12, 12)
        Me.BTN_Code_Manager.Name = "BTN_Code_Manager"
        Me.BTN_Code_Manager.Size = New System.Drawing.Size(254, 178)
        Me.BTN_Code_Manager.TabIndex = 16
        Me.BTN_Code_Manager.Text = "Code 관리"
        Me.BTN_Code_Manager.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BTN_Code_Manager.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.BTN_Code_Manager.UseVisualStyleBackColor = False
        '
        'Subform
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.BTN_Cal_Object)
        Me.Controls.Add(Me.BTN_uploadPG)
        Me.Controls.Add(Me.BTN_connection)
        Me.Controls.Add(Me.BTN_Asset_Inspection)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.BTN_Close)
        Me.Controls.Add(Me.BTN_Asset_History)
        Me.Controls.Add(Me.BTN_Asset_Manager2)
        Me.Controls.Add(Me.BTN_Asset_Manager)
        Me.Controls.Add(Me.BTN_Asset_ADD)
        Me.Controls.Add(Me.BTN_Code_Manager)
        Me.Name = "Subform"
        Me.Text = "Home"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BTN_Code_Manager As Button
    Friend WithEvents BTN_Close As Button
    Friend WithEvents BTN_Asset_History As Button
    Friend WithEvents BTN_Asset_Manager As Button
    Friend WithEvents BTN_Asset_ADD As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents BTN_Asset_Manager2 As Button
    Friend WithEvents BTN_Asset_Inspection As Button
    Friend WithEvents BTN_connection As Button
    Friend WithEvents BTN_uploadPG As Button
    Friend WithEvents BTN_Cal_Object As Button
End Class

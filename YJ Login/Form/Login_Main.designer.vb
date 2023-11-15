<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
<Global.System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1726")> _
Partial Class Login_Main
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
    Friend WithEvents LogoPictureBox As System.Windows.Forms.PictureBox
    Friend WithEvents UsernameLabel As System.Windows.Forms.Label
    Friend WithEvents PasswordLabel As System.Windows.Forms.Label
    Friend WithEvents USER_ID As System.Windows.Forms.TextBox
    Friend WithEvents USER_PASS As System.Windows.Forms.TextBox
    Friend WithEvents OK As System.Windows.Forms.Button

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Login_Main))
        Me.UsernameLabel = New System.Windows.Forms.Label()
        Me.PasswordLabel = New System.Windows.Forms.Label()
        Me.USER_ID = New System.Windows.Forms.TextBox()
        Me.USER_PASS = New System.Windows.Forms.TextBox()
        Me.OK = New System.Windows.Forms.Button()
        Me.LogoPictureBox = New System.Windows.Forms.PictureBox()
        Me.CB_PASS_SAVE = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cb_RunProgram = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UsernameLabel
        '
        Me.UsernameLabel.Location = New System.Drawing.Point(23, 133)
        Me.UsernameLabel.Name = "UsernameLabel"
        Me.UsernameLabel.Size = New System.Drawing.Size(220, 23)
        Me.UsernameLabel.TabIndex = 0
        Me.UsernameLabel.Text = "사용자 이름(&U)"
        Me.UsernameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PasswordLabel
        '
        Me.PasswordLabel.Location = New System.Drawing.Point(23, 177)
        Me.PasswordLabel.Name = "PasswordLabel"
        Me.PasswordLabel.Size = New System.Drawing.Size(220, 23)
        Me.PasswordLabel.TabIndex = 2
        Me.PasswordLabel.Text = "암호(&P)"
        Me.PasswordLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'USER_ID
        '
        Me.USER_ID.Location = New System.Drawing.Point(25, 153)
        Me.USER_ID.Name = "USER_ID"
        Me.USER_ID.Size = New System.Drawing.Size(220, 21)
        Me.USER_ID.TabIndex = 1
        '
        'USER_PASS
        '
        Me.USER_PASS.Location = New System.Drawing.Point(25, 197)
        Me.USER_PASS.Name = "USER_PASS"
        Me.USER_PASS.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.USER_PASS.Size = New System.Drawing.Size(220, 21)
        Me.USER_PASS.TabIndex = 3
        '
        'OK
        '
        Me.OK.Image = Global.YJ_Login.My.Resources.Resources._102_16
        Me.OK.Location = New System.Drawing.Point(273, 151)
        Me.OK.Name = "OK"
        Me.OK.Size = New System.Drawing.Size(94, 67)
        Me.OK.TabIndex = 4
        Me.OK.Text = "확인(&O)"
        Me.OK.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.OK.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.OK.UseVisualStyleBackColor = True
        '
        'LogoPictureBox
        '
        Me.LogoPictureBox.Image = Global.YJ_Login.My.Resources.Resources.YUJIN_Logo
        Me.LogoPictureBox.Location = New System.Drawing.Point(12, 12)
        Me.LogoPictureBox.Name = "LogoPictureBox"
        Me.LogoPictureBox.Size = New System.Drawing.Size(382, 63)
        Me.LogoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.LogoPictureBox.TabIndex = 0
        Me.LogoPictureBox.TabStop = False
        '
        'CB_PASS_SAVE
        '
        Me.CB_PASS_SAVE.AutoSize = True
        Me.CB_PASS_SAVE.Location = New System.Drawing.Point(25, 224)
        Me.CB_PASS_SAVE.Name = "CB_PASS_SAVE"
        Me.CB_PASS_SAVE.Size = New System.Drawing.Size(100, 16)
        Me.CB_PASS_SAVE.TabIndex = 6
        Me.CB_PASS_SAVE.Text = "비밀번호 저장"
        Me.CB_PASS_SAVE.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(23, 89)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(220, 23)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "접속 위치(&S)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cb_RunProgram
        '
        Me.cb_RunProgram.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_RunProgram.FormattingEnabled = True
        Me.cb_RunProgram.Items.AddRange(New Object() {"Repair System", "오삽방지 시스템"})
        Me.cb_RunProgram.Location = New System.Drawing.Point(25, 110)
        Me.cb_RunProgram.Name = "cb_RunProgram"
        Me.cb_RunProgram.Size = New System.Drawing.Size(220, 20)
        Me.cb_RunProgram.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(227, 229)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(168, 12)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "※ 현재 Login 기능 사용 안됨."
        '
        'Login_Main
        '
        Me.AcceptButton = Me.OK
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(407, 250)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cb_RunProgram)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CB_PASS_SAVE)
        Me.Controls.Add(Me.OK)
        Me.Controls.Add(Me.USER_PASS)
        Me.Controls.Add(Me.USER_ID)
        Me.Controls.Add(Me.PasswordLabel)
        Me.Controls.Add(Me.UsernameLabel)
        Me.Controls.Add(Me.LogoPictureBox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Login_Main"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Log-in"
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CB_PASS_SAVE As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cb_RunProgram As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As Label
End Class

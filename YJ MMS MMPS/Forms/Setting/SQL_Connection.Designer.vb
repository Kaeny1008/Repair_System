<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SQL_Connection
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
        Me.LB_IP = New System.Windows.Forms.Label()
        Me.TB_IP = New System.Windows.Forms.TextBox()
        Me.TB_PORT = New System.Windows.Forms.TextBox()
        Me.LB_PORT = New System.Windows.Forms.Label()
        Me.TB_ID = New System.Windows.Forms.TextBox()
        Me.LB_ID = New System.Windows.Forms.Label()
        Me.TB_PW = New System.Windows.Forms.TextBox()
        Me.LB_PW = New System.Windows.Forms.Label()
        Me.BTN_SAVE = New System.Windows.Forms.Button()
        Me.UPDATE_CHECK = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'LB_IP
        '
        Me.LB_IP.BackColor = System.Drawing.Color.LightSteelBlue
        Me.LB_IP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LB_IP.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold)
        Me.LB_IP.ForeColor = System.Drawing.Color.White
        Me.LB_IP.Location = New System.Drawing.Point(12, 8)
        Me.LB_IP.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.LB_IP.Name = "LB_IP"
        Me.LB_IP.Size = New System.Drawing.Size(191, 26)
        Me.LB_IP.TabIndex = 0
        Me.LB_IP.Text = "SQL IP :"
        Me.LB_IP.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TB_IP
        '
        Me.TB_IP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_IP.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold)
        Me.TB_IP.Location = New System.Drawing.Point(203, 8)
        Me.TB_IP.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_IP.Name = "TB_IP"
        Me.TB_IP.Size = New System.Drawing.Size(171, 26)
        Me.TB_IP.TabIndex = 1
        '
        'TB_PORT
        '
        Me.TB_PORT.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_PORT.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold)
        Me.TB_PORT.Location = New System.Drawing.Point(203, 40)
        Me.TB_PORT.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_PORT.Name = "TB_PORT"
        Me.TB_PORT.Size = New System.Drawing.Size(171, 26)
        Me.TB_PORT.TabIndex = 3
        '
        'LB_PORT
        '
        Me.LB_PORT.BackColor = System.Drawing.Color.LightSteelBlue
        Me.LB_PORT.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LB_PORT.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold)
        Me.LB_PORT.ForeColor = System.Drawing.Color.White
        Me.LB_PORT.Location = New System.Drawing.Point(12, 40)
        Me.LB_PORT.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.LB_PORT.Name = "LB_PORT"
        Me.LB_PORT.Size = New System.Drawing.Size(191, 26)
        Me.LB_PORT.TabIndex = 2
        Me.LB_PORT.Text = "SQL Port :"
        Me.LB_PORT.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TB_ID
        '
        Me.TB_ID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_ID.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold)
        Me.TB_ID.Location = New System.Drawing.Point(203, 72)
        Me.TB_ID.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_ID.Name = "TB_ID"
        Me.TB_ID.Size = New System.Drawing.Size(171, 26)
        Me.TB_ID.TabIndex = 5
        '
        'LB_ID
        '
        Me.LB_ID.BackColor = System.Drawing.Color.LightSteelBlue
        Me.LB_ID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LB_ID.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold)
        Me.LB_ID.ForeColor = System.Drawing.Color.White
        Me.LB_ID.Location = New System.Drawing.Point(12, 72)
        Me.LB_ID.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.LB_ID.Name = "LB_ID"
        Me.LB_ID.Size = New System.Drawing.Size(191, 26)
        Me.LB_ID.TabIndex = 4
        Me.LB_ID.Text = "SQL ID :"
        Me.LB_ID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TB_PW
        '
        Me.TB_PW.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_PW.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold)
        Me.TB_PW.Location = New System.Drawing.Point(203, 104)
        Me.TB_PW.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TB_PW.Name = "TB_PW"
        Me.TB_PW.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TB_PW.Size = New System.Drawing.Size(171, 26)
        Me.TB_PW.TabIndex = 7
        '
        'LB_PW
        '
        Me.LB_PW.BackColor = System.Drawing.Color.LightSteelBlue
        Me.LB_PW.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LB_PW.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold)
        Me.LB_PW.ForeColor = System.Drawing.Color.White
        Me.LB_PW.Location = New System.Drawing.Point(12, 104)
        Me.LB_PW.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.LB_PW.Name = "LB_PW"
        Me.LB_PW.Size = New System.Drawing.Size(191, 26)
        Me.LB_PW.TabIndex = 6
        Me.LB_PW.Text = "SQL P/W :"
        Me.LB_PW.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BTN_SAVE
        '
        Me.BTN_SAVE.Font = New System.Drawing.Font("굴림", 15.75!, System.Drawing.FontStyle.Bold)
        Me.BTN_SAVE.Location = New System.Drawing.Point(377, 79)
        Me.BTN_SAVE.Name = "BTN_SAVE"
        Me.BTN_SAVE.Size = New System.Drawing.Size(102, 83)
        Me.BTN_SAVE.TabIndex = 8
        Me.BTN_SAVE.Text = "저장 및 " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "닫기"
        Me.BTN_SAVE.UseVisualStyleBackColor = True
        '
        'UPDATE_CHECK
        '
        Me.UPDATE_CHECK.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.UPDATE_CHECK.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold)
        Me.UPDATE_CHECK.Location = New System.Drawing.Point(203, 136)
        Me.UPDATE_CHECK.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.UPDATE_CHECK.Name = "UPDATE_CHECK"
        Me.UPDATE_CHECK.Size = New System.Drawing.Size(171, 26)
        Me.UPDATE_CHECK.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(12, 136)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(191, 26)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Update 확인 주기(초)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'SQL_Connection
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.LightSteelBlue
        Me.ClientSize = New System.Drawing.Size(484, 174)
        Me.Controls.Add(Me.UPDATE_CHECK)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BTN_SAVE)
        Me.Controls.Add(Me.TB_PW)
        Me.Controls.Add(Me.LB_PW)
        Me.Controls.Add(Me.TB_ID)
        Me.Controls.Add(Me.LB_ID)
        Me.Controls.Add(Me.TB_PORT)
        Me.Controls.Add(Me.LB_PORT)
        Me.Controls.Add(Me.TB_IP)
        Me.Controls.Add(Me.LB_IP)
        Me.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SQL_Connection"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SQL 접속정보 설정"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LB_IP As System.Windows.Forms.Label
    Friend WithEvents TB_IP As System.Windows.Forms.TextBox
    Friend WithEvents TB_PORT As System.Windows.Forms.TextBox
    Friend WithEvents LB_PORT As System.Windows.Forms.Label
    Friend WithEvents TB_ID As System.Windows.Forms.TextBox
    Friend WithEvents LB_ID As System.Windows.Forms.Label
    Friend WithEvents TB_PW As System.Windows.Forms.TextBox
    Friend WithEvents LB_PW As System.Windows.Forms.Label
    Friend WithEvents BTN_SAVE As System.Windows.Forms.Button
    Friend WithEvents UPDATE_CHECK As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class

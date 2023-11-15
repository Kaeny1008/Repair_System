<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProgressbarEffect
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
        Me.PBar = New System.Windows.Forms.ProgressBar
        Me.PbarPercent = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'PBar
        '
        Me.PBar.Location = New System.Drawing.Point(12, 25)
        Me.PBar.Name = "PBar"
        Me.PBar.Size = New System.Drawing.Size(279, 23)
        Me.PBar.TabIndex = 50
        '
        'PbarPercent
        '
        Me.PbarPercent.BackColor = System.Drawing.Color.Transparent
        Me.PbarPercent.Location = New System.Drawing.Point(12, 7)
        Me.PbarPercent.Name = "PbarPercent"
        Me.PbarPercent.Size = New System.Drawing.Size(279, 14)
        Me.PbarPercent.TabIndex = 51
        Me.PbarPercent.Text = "불러오는중"
        Me.PbarPercent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Location = New System.Drawing.Point(89, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(125, 12)
        Me.Label1.TabIndex = 52
        Me.Label1.Text = "잠시만 기다려 주세요."
        '
        'ProgressbarEffect
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(225, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(303, 81)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PBar)
        Me.Controls.Add(Me.PbarPercent)
        Me.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "ProgressbarEffect"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ShowEffect"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PBar As System.Windows.Forms.ProgressBar
    Friend WithEvents PbarPercent As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class

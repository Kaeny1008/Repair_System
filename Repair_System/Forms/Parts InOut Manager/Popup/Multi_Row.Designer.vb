<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Multi_Row
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.TB_Row_Count = New System.Windows.Forms.TextBox
        Me.BTN_Insert = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.SteelBlue
        Me.Label1.Font = New System.Drawing.Font("굴림", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(12, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(430, 30)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "추가하려는 행 수를 입력한 뒤 추가 버튼을 눌러주세요."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TB_Row_Count
        '
        Me.TB_Row_Count.Font = New System.Drawing.Font("굴림", 12.0!)
        Me.TB_Row_Count.Location = New System.Drawing.Point(15, 54)
        Me.TB_Row_Count.Name = "TB_Row_Count"
        Me.TB_Row_Count.Size = New System.Drawing.Size(333, 26)
        Me.TB_Row_Count.TabIndex = 1
        '
        'BTN_Insert
        '
        Me.BTN_Insert.Font = New System.Drawing.Font("굴림", 12.0!)
        Me.BTN_Insert.Location = New System.Drawing.Point(354, 54)
        Me.BTN_Insert.Name = "BTN_Insert"
        Me.BTN_Insert.Size = New System.Drawing.Size(88, 26)
        Me.BTN_Insert.TabIndex = 2
        Me.BTN_Insert.Text = "추가"
        Me.BTN_Insert.UseVisualStyleBackColor = True
        '
        'Multi_Row
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(472, 93)
        Me.Controls.Add(Me.BTN_Insert)
        Me.Controls.Add(Me.TB_Row_Count)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Multi_Row"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "다중 행 추가"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TB_Row_Count As System.Windows.Forms.TextBox
    Friend WithEvents BTN_Insert As System.Windows.Forms.Button
End Class

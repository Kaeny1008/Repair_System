Public Class frm_M_Second_Popup

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim OpenFileDialog1 As New OpenFileDialog

        OpenFileDialog1.Filter = "Excel (*.xls, *.xlsx)|*.xls;*.xlsx"
        OpenFileDialog1.InitialDirectory = "D:\1. Workplace\5. 고객사 자료\18. 삼성전자(Repair)\##1. OMS Upload\3. OMS Excel Download\2. Return Type"
        OpenFileDialog1.ShowDialog()

        If Not OpenFileDialog1.FileName = String.Empty Then
            TB_File_Path.Text = OpenFileDialog1.FileName
        Else
            Exit Sub
        End If

        OpenFileDialog1.Dispose()

        'Dim slip_no, product_code, lot_no, lot_type, wafer_qty, chip_qty, confirm, ww, issue_date, reg_date, confirm_date As Integer
        frm_Module_In_Add.lot_no = TextBox1.Text
        frm_Module_In_Add.return_type = TextBox2.Text
        frm_Module_In_Add.start_row = TextBox3.Text
        frm_Module_In_Add.confirm_date = TextBox4.Text
        frm_Module_In_Add.open_File_Path = TB_File_Path.Text

        frm_Module_In_Add.File_Open_Start2()

        Me.Dispose()

    End Sub

    Private Sub Second_Popup_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox1.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Lot No_2", 2)
        TextBox2.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Return Type", 4)
        TextBox3.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Start Row_2", 3)
        TextBox4.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Confirm Date_2", 12)

    End Sub

    Private Sub Second_Popup_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Lot No_2", TextBox1.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Return Type", TextBox2.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Start Row_2", TextBox3.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Confirm Date_2", TextBox4.Text)

    End Sub
End Class
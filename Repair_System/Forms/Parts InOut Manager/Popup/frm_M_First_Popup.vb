Public Class frm_M_First_Popup

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim OpenFileDialog1 As New OpenFileDialog

        OpenFileDialog1.Filter = "Excel (*.xls, *.xlsx)|*.xls;*.xlsx"
        OpenFileDialog1.InitialDirectory = "D:\1. Workplace\5. 고객사 자료\18. 삼성전자(Repair)\##1. OMS Upload\3. OMS Excel Download\1. 상세"
        OpenFileDialog1.ShowDialog()

        If Not OpenFileDialog1.FileName = String.Empty Then
            TB_File_Path.Text = OpenFileDialog1.FileName
        Else
            Exit Sub
        End If

        OpenFileDialog1.Dispose()

        'Dim slip_no, product_code, lot_no, lot_type, wafer_qty, chip_qty, confirm, ww, issue_date, reg_date, confirm_date As Integer
        frm_Module_In_Add.slip_no = TextBox1.Text
        frm_Module_In_Add.product_code = TextBox2.Text
        frm_Module_In_Add.lot_no = TextBox3.Text
        frm_Module_In_Add.lot_type = TextBox4.Text
        frm_Module_In_Add.wafer_qty = TextBox5.Text
        frm_Module_In_Add.chip_qty = TextBox6.Text
        frm_Module_In_Add.confirm = TextBox7.Text
        frm_Module_In_Add.ww = TextBox8.Text
        frm_Module_In_Add.work_step = TextBox13.Text
        frm_Module_In_Add.issue_date = TextBox9.Text
        frm_Module_In_Add.reg_date = TextBox10.Text
        frm_Module_In_Add.confirm_date = TextBox11.Text
        frm_Module_In_Add.start_row = TextBox12.Text
        frm_Module_In_Add.open_File_Path = TB_File_Path.Text

        frm_Module_In_Add.File_Open_Start1()

        Me.Dispose()

    End Sub

    Private Sub First_Popup_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox1.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Slip No", 2)
        TextBox2.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Product", 3)
        TextBox3.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Lot No", 4)
        TextBox4.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Lot Type", 5)
        TextBox5.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Wafer Qty", 6)
        TextBox6.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Chip Qty", 7)
        TextBox7.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Confirm", 8)
        TextBox8.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "WW", 9)
        TextBox13.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Step", 10)
        TextBox9.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Issue Date", 11)
        TextBox10.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Reg Date", 12)
        TextBox11.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Confirm Date", 13)
        TextBox12.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Start Row", 3)

    End Sub

    Private Sub First_Popup_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Slip No", TextBox1.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Product", TextBox2.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Lot No", TextBox3.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Lot Type", TextBox4.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Wafer Qty", TextBox5.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Chip Qty", TextBox6.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Confirm", TextBox7.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "WW", TextBox8.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Issue Date", TextBox9.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Reg Date", TextBox10.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Confirm Date", TextBox11.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Start Row", TextBox12.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Step", TextBox13.Text)

    End Sub
End Class
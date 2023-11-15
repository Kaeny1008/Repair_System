Public Class frm_M_Third_Popup

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim OpenFileDialog1 As New OpenFileDialog

        OpenFileDialog1.Filter = "Excel (*.xls, *.xlsx)|*.xls;*.xlsx"
        OpenFileDialog1.InitialDirectory = "D:\1. Workplace\5. 고객사 자료\18. 삼성전자(Repair)\##1. OMS Upload\4. 우익 수신"
        OpenFileDialog1.ShowDialog()

        If Not OpenFileDialog1.FileName = String.Empty Then
            TB_File_Path.Text = OpenFileDialog1.FileName
        Else
            Exit Sub
        End If

        OpenFileDialog1.Dispose()

        'Dim slip_no, product_code, lot_no, lot_type, wafer_qty, chip_qty, confirm, ww, issue_date, reg_date, confirm_date As Integer
        frm_Module_In_Add.lot_no = TextBox1.Text
        frm_Module_In_Add.fab_line = TextBox2.Text
        frm_Module_In_Add.lot_option = TextBox3.Text
        frm_Module_In_Add.start_row = TextBox4.Text
        frm_Module_In_Add.open_File_Path = TB_File_Path.Text

        frm_Module_In_Add.File_Open_Start3()

        frm_M_Third_Popup_FormClosing(Nothing, Nothing)

        Me.Dispose()

    End Sub

    Private Sub frm_M_Third_Popup_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox1.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Lot No_3", 3)
        TextBox2.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Fab Line", 11)
        TextBox3.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Lot Option", 12)
        TextBox4.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Start Row_4", 3)

    End Sub

    Private Sub frm_M_Third_Popup_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Lot No_3", TextBox1.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Fab Line", TextBox2.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Lot Option", TextBox3.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\OMS Excel Setting", "Start Row_4", TextBox4.Text)

    End Sub
End Class
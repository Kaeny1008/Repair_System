Public Class frm_P_First_Popup

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim OpenFileDialog1 As New OpenFileDialog

        OpenFileDialog1.Filter = "Excel (*.xls, *.xlsx)|*.xls;*.xlsx"
        OpenFileDialog1.ShowDialog()

        If Not OpenFileDialog1.FileName = String.Empty Then
            TB_File_Path.Text = OpenFileDialog1.FileName
        Else
            Exit Sub
        End If

        OpenFileDialog1.Dispose()

        frm_Material_In_Add.slip_no = CInt(tb_SlipNo.Text)
        frm_Material_In_Add.part_no = CInt(tb_PartNo.Text)
        frm_Material_In_Add.lot_no = CInt(tb_LotNo.Text)
        frm_Material_In_Add.qty = CInt(tb_Qty.Text)
        frm_Material_In_Add.start_row = CInt(tb_FirstRow.Text)
        frm_Material_In_Add.open_File_Path = TB_File_Path.Text

        frm_Material_In_Add.File_Open_Start()

        Me.Dispose()

    End Sub

    Private Sub First_Popup_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        tb_SlipNo.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Parts Excel Setting", "Slip No", 2)
        tb_PartNo.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Parts Excel Setting", "Part No", 3)
        tb_LotNo.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Parts Excel Setting", "Lot No", 4)
        tb_Qty.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Parts Excel Setting", "Qty", 7)
        tb_FirstRow.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Parts Excel Setting", "Start Row", 3)

    End Sub

    Private Sub First_Popup_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Parts Excel Setting", "Slip No", tb_SlipNo.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\Parts Excel Setting", "Part No", tb_PartNo.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\Parts Excel Setting", "Lot No", tb_LotNo.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\Parts Excel Setting", "Qty", tb_Qty.Text)
        registryEdit.WriteRegKey("Software\Yujin\Repair System\Parts Excel Setting", "Start Row", tb_FirstRow.Text)

    End Sub

    Private Sub tb_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_SlipNo.KeyDown,
        tb_PartNo.KeyDown,
        tb_LotNo.KeyDown,
        tb_Qty.KeyDown,
        tb_FirstRow.KeyDown

        Select Case e.KeyCode
            Case 13
            Case 8
            Case 96 To 105, 48 To 57
            Case Else
                e.SuppressKeyPress = True
        End Select

    End Sub
End Class
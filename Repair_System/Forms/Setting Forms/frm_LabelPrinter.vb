'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-07-27
'##### 수정일자 : 2023-07-27
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 실명제 관련 각종 라벨을 발행한다.

'############################################################################################################
'############################################################################################################

Imports System.IO
Imports System.Text

Public Class frm_LabelPrinter

    Private WithEvents ComPort As New System.IO.Ports.SerialPort

    Private Sub frm_LabelPrinter_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Setting_LOAD()

    End Sub

    Private Sub Setting_LOAD()

        cb_Cable.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "COM/LPT", "COM")
        cb_PrinterList.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Printer Name", "")
        TB_Port.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Port", 1)
        TB_LEFT_Loc.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Width", 0)
        TB_TOP_Loc.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Height", 0)
        TB_BaudRate.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "BaudRate", 9600)
        TB_DataBits.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "DataBits", 8)
        TB_Parity.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Parity", 0)
        TB_StopBits.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "StopBits", 1)

    End Sub

    Private Sub TB_Port_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Port.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Port", TB_Port.Text)

    End Sub

    Private Sub TB_Width_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_LEFT_Loc.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Width", TB_LEFT_Loc.Text)

    End Sub

    Private Sub TB_Height_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_TOP_Loc.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Height", TB_TOP_Loc.Text)

    End Sub

    Private Sub TB_BaudRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_BaudRate.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "BaudRate", TB_BaudRate.Text)

    End Sub

    Private Sub TB_DataBits_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_DataBits.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "DataBits", TB_DataBits.Text)

    End Sub

    Private Sub TB_Parity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Parity.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Parity", TB_Parity.Text)

    End Sub

    Private Sub TB_StopBits_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_StopBits.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "StopBits", TB_StopBits.Text)

    End Sub

    Private Sub cb_PrinterList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_PrinterList.SelectedIndexChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Printer Name", cb_PrinterList.Text)

    End Sub

    Private Sub cb_Cable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_Cable.SelectedIndexChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "COM/LPT", cb_Cable.Text)

        If cb_Cable.Text = "USB" Then
            PrinterListLoad()
        End If

    End Sub

    Private Sub PrinterListLoad()

        cb_PrinterList.Items.Clear()

        For i As Integer = 0 To Printing.PrinterSettings.InstalledPrinters.Count - 1
            Dim Printers As String = Printing.PrinterSettings.InstalledPrinters.Item(i)
            cb_PrinterList.Items.Add(Printers)
        Next

        cb_PrinterList.Sorted = True

    End Sub

    Private Sub pb_PMIC_Label_Click(sender As Object, e As EventArgs) Handles pb_PMIC_Label.Click

        RadioButton1.Checked = True

    End Sub

    Private Sub pb_RCD_Label_Click(sender As Object, e As EventArgs) Handles pb_RCD_Label.Click

        RadioButton2.Checked = True

    End Sub

    Private Sub pb_HeatCount1_Label_Click(sender As Object, e As EventArgs) Handles pb_HeatCount1_Label.Click

        RadioButton3.Checked = True

    End Sub

    Private Sub pb_HeatCount2_Label_Click(sender As Object, e As EventArgs) Handles pb_HeatCount2_Label.Click

        RadioButton4.Checked = True

    End Sub

    Private Sub btn_LabelPrint_Click(sender As Object, e As EventArgs) Handles btn_LabelPrint.Click

        If MsgBox("선택된 라벨을 발행 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "각종 라벨 발행") = MsgBoxResult.No Then Exit Sub

        If RadioButton1.Checked Then
            PMIC_Label_Print()
        ElseIf RadioButton2.Checked Then
            RCD_Label_Print()
        ElseIf RadioButton3.Checked Then
            HeatCount1_Label_Print()
        ElseIf RadioButton4.Checked Then
            HeatCount2_Label_Print()
        ElseIf RadioButton5.Checked Then
            TestFail_Label_Print()
        ElseIf RadioButton6.Checked Then
            MapOut_Label_Print()
        ElseIf RadioButton7.Checked Then
            PMIC_Reg_FAIL_Label_Print()
        ElseIf RadioButton8.Checked Then
            None_Label_Print()
        ElseIf RadioButton9.Checked Then
            Process1_Label_Print()
        ElseIf RadioButton10.Checked Then
            Process2_Label_Print()
        End If

        MsgBox("라벨 발행 완료." & vbCrLf & "인쇄된 라벨을 확인하여 주십시오.", MsgBoxStyle.Information, "각종 라벨 발행")

    End Sub

    Private Sub HeatCount2_Label_Print()

        If cb_Cable.Text = "USB" Then
            Dim p As New PrintRaw
            Dim s As New StringBuilder
            s.AppendLine("^XZ~JA^XZ")
            s.AppendLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            s.AppendLine("^MD15^PR2,2,2") '진하기
            s.AppendLine("^FO0000,0000^GB684,155,2,B,0^FS")
            s.AppendLine("^FO0000,0055^GB684,000,2,B,0^FS")
            s.AppendLine("^FO0114,0000^GB000,155,2,B,0^FS")
            s.AppendLine("^FO0228,0000^GB000,155,2,B,0^FS")
            s.AppendLine("^FO0342,0000^GB000,155,2,B,0^FS")
            s.AppendLine("^FO0456,0000^GB000,155,2,B,0^FS")
            s.AppendLine("^FO0570,0000^GB000,155,2,B,0^FS")
            s.AppendLine("^FO0040,0012^A0,40,30^FD7M^FS")
            s.AppendLine("^FO0154,0012^A0,40,30^FD8M^FS")
            s.AppendLine("^FO0268,0012^A0,40,30^FD9M^FS")
            s.AppendLine("^FO0372,0012^A0,40,30^FD10M^FS")
            s.AppendLine("^FO0486,0012^A0,40,30^FD11M^FS")
            s.AppendLine("^FO0600,0010^A0,40,30^FD12M^FS")
            s.AppendLine("^FO022,0062^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO136,0062^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO250,0062^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO364,0062^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO478,0062^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO589,0060^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO052,0092^A0,30,30^FD3^FS")
            s.AppendLine("^FO166,0092^A0,30,30^FD3^FS")
            s.AppendLine("^FO280,0092^A0,30,30^FD3^FS")
            s.AppendLine("^FO394,0092^A0,30,30^FD3^FS")
            s.AppendLine("^FO508,0092^A0,30,30^FD3^FS")
            s.AppendLine("^FO622,0090^A0,30,30^FD3^FS")
            s.AppendLine("^FO022,0122^A0,30,30^FD4     5^FS")
            s.AppendLine("^FO136,0122^A0,30,30^FD4     5^FS")
            s.AppendLine("^FO250,0122^A0,30,30^FD4     5^FS")
            s.AppendLine("^FO364,0122^A0,30,30^FD4     5^FS")
            s.AppendLine("^FO478,0122^A0,30,30^FD4     5^FS")
            s.AppendLine("^FO589,0120^A0,30,30^FD4     5^FS")
            s.AppendLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            p.Print(s, cb_PrinterList.Text)
        Else
            Dim swFile As StreamWriter
            swFile = New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))
            swFile.WriteLine("^XZ~JA^XZ")
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15^PR2,2,2") '진하기
            swFile.WriteLine("^FO0000,0000^GB462,108,2,B,0^FS")
            swFile.WriteLine("^FO0000,0028^GB462,000,2,B,0^FS")
            swFile.WriteLine("^FO0077,0000^GB000,108,2,B,0^FS")
            swFile.WriteLine("^FO0154,0000^GB000,108,2,B,0^FS")
            swFile.WriteLine("^FO0231,0000^GB000,108,2,B,0^FS")
            swFile.WriteLine("^FO0308,0000^GB000,108,2,B,0^FS")
            swFile.WriteLine("^FO0385,0000^GB000,108,2,B,0^FS")
            swFile.WriteLine("^FO0024,0004^A0,30,20^FD7M^FS")
            swFile.WriteLine("^FO0101,0004^A0,30,20^FD8M^FS")
            swFile.WriteLine("^FO0178,0004^A0,30,20^FD9M^FS")
            swFile.WriteLine("^FO0255,0004^A0,30,20^FD10M^FS")
            swFile.WriteLine("^FO0332,0004^A0,30,20^FD11M^FS")
            swFile.WriteLine("^FO0409,0004^A0,30,20^FD12M^FS")
            swFile.WriteLine("^FO012,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO089,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO166,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO243,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO320,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO397,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO032,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO109,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO186,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO263,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO340,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO417,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO012,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^FO089,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^FO166,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^FO243,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^FO320,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^FO397,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", cb_Cable.Text & TB_Port.Text)
            Else

                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While
                sr.Close()
                ComPort.Close()
            End If
        End If
        File.Delete(Application.StartupPath & "\print.txt")

    End Sub

    Private Sub HeatCount1_Label_Print()

        If cb_Cable.Text = "USB" Then
            Dim p As New PrintRaw
            Dim s As New StringBuilder
            s.AppendLine("^XZ~JA^XZ")
            s.AppendLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            s.AppendLine("^MD15^PR2,2,2") '진하기
            s.AppendLine("^FO0000,0000^GB684,155,2,B,0^FS")
            s.AppendLine("^FO0000,0055^GB684,000,2,B,0^FS")
            s.AppendLine("^FO0114,0000^GB000,155,2,B,0^FS")
            s.AppendLine("^FO0228,0000^GB000,155,2,B,0^FS")
            s.AppendLine("^FO0342,0000^GB000,155,2,B,0^FS")
            s.AppendLine("^FO0456,0000^GB000,155,2,B,0^FS")
            s.AppendLine("^FO0570,0000^GB000,155,2,B,0^FS")
            s.AppendLine("^FO0040,0012^A0,40,30^FD1M^FS")
            s.AppendLine("^FO0154,0012^A0,40,30^FD2M^FS")
            s.AppendLine("^FO0268,0012^A0,40,30^FD3M^FS")
            s.AppendLine("^FO0382,0012^A0,40,30^FD4M^FS")
            s.AppendLine("^FO0496,0012^A0,40,30^FD5M^FS")
            s.AppendLine("^FO0610,0010^A0,40,30^FD6M^FS")
            s.AppendLine("^FO022,0062^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO136,0062^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO250,0062^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO364,0062^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO478,0062^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO589,0060^A0,30,30^FD1     2^FS")
            s.AppendLine("^FO052,0092^A0,30,30^FD3^FS")
            s.AppendLine("^FO166,0092^A0,30,30^FD3^FS")
            s.AppendLine("^FO280,0092^A0,30,30^FD3^FS")
            s.AppendLine("^FO394,0092^A0,30,30^FD3^FS")
            s.AppendLine("^FO508,0092^A0,30,30^FD3^FS")
            s.AppendLine("^FO622,0090^A0,30,30^FD3^FS")
            s.AppendLine("^FO022,0122^A0,30,30^FD4     5^FS")
            s.AppendLine("^FO136,0122^A0,30,30^FD4     5^FS")
            s.AppendLine("^FO250,0122^A0,30,30^FD4     5^FS")
            s.AppendLine("^FO364,0122^A0,30,30^FD4     5^FS")
            s.AppendLine("^FO478,0122^A0,30,30^FD4     5^FS")
            s.AppendLine("^FO589,0120^A0,30,30^FD4     5^FS")
            s.AppendLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            p.Print(s, cb_PrinterList.Text)
        Else
            Dim swFile As StreamWriter
            swFile = New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))
            swFile.WriteLine("^XZ~JA^XZ")
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15^PR2,2,2") '진하기
            swFile.WriteLine("^FO0000,0000^GB462,108,2,B,0^FS")
            swFile.WriteLine("^FO0000,0028^GB462,000,2,B,0^FS")
            swFile.WriteLine("^FO0077,0000^GB000,108,2,B,0^FS")
            swFile.WriteLine("^FO0154,0000^GB000,108,2,B,0^FS")
            swFile.WriteLine("^FO0231,0000^GB000,108,2,B,0^FS")
            swFile.WriteLine("^FO0308,0000^GB000,108,2,B,0^FS")
            swFile.WriteLine("^FO0385,0000^GB000,108,2,B,0^FS")
            swFile.WriteLine("^FO0024,0004^A0,30,20^FD1M^FS")
            swFile.WriteLine("^FO0101,0004^A0,30,20^FD2M^FS")
            swFile.WriteLine("^FO0178,0004^A0,30,20^FD3M^FS")
            swFile.WriteLine("^FO0255,0004^A0,30,20^FD4M^FS")
            swFile.WriteLine("^FO0332,0004^A0,30,20^FD5M^FS")
            swFile.WriteLine("^FO0409,0004^A0,30,20^FD6M^FS")
            swFile.WriteLine("^FO012,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO089,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO166,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO243,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO320,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO397,0040^A0,20,15^FD1         2^FS")
            swFile.WriteLine("^FO032,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO109,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO186,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO263,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO340,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO417,0060^A0,20,15^FD3^FS")
            swFile.WriteLine("^FO012,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^FO089,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^FO166,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^FO243,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^FO320,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^FO397,0080^A0,20,15^FD4         5^FS")
            swFile.WriteLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", cb_Cable.Text & TB_Port.Text)
            Else

                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While
                sr.Close()
                ComPort.Close()
            End If
        End If
        File.Delete(Application.StartupPath & "\print.txt")

    End Sub

    Private Sub RCD_Label_Print()

        Try
            Dim swFile As StreamWriter =
                New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))

            swFile.WriteLine("^XZ~JA^XZ")
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15^PR2,2,2") '진하기
            swFile.WriteLine("^FO0000,0000^GB513,155,2,B,0^FS")
            swFile.WriteLine("^FO0000,0055^GB513,000,2,B,0^FS")
            swFile.WriteLine("^FO0000,0090^GB513,000,2,B,0^FS")
            swFile.WriteLine("^FO0171,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0342,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0045,0015^A0,40,30^FDDetach^FS")
            swFile.WriteLine("^FO0007,0070^A0,15,15^FDDate                  /^FS")
            swFile.WriteLine("^FO0007,0100^A0,15,15^FDName^FS")
            swFile.WriteLine("^FO0205,0015^A0,40,30^FDDressing^FS")
            swFile.WriteLine("^FO0178,0070^A0,15,15^FDDate                  /^FS")
            swFile.WriteLine("^FO0178,0100^A0,15,15^FDName^FS")
            swFile.WriteLine("^FO0360,0015^A0,40,30^FDDressing 2^FS")
            swFile.WriteLine("^FO0349,0070^A0,15,15^FDDate                  /^FS")
            swFile.WriteLine("^FO0349,0100^A0,15,15^FDName^FS")
            swFile.WriteLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", "LPT" & TB_Port.Text)
            ElseIf cb_Cable.Text = "COM" Then
                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While

                sr.Close()
                ComPort.Close()
            ElseIf cb_Cable.Text = "USB" Then
                Dim p As New PrintRaw
                Dim s As New StringBuilder
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    s.AppendLine(str)
                End While

                sr.Close()
                p.Print(s, cb_PrinterList.Text)
            End If
            File.Delete(Application.StartupPath & "\print.txt")
        Catch ex As Exception
            File.Delete(Application.StartupPath & "\print.txt")
            MsgBox(ex.Message, MsgBoxStyle.Critical, "각종 라벨 발행")
        End Try

    End Sub

    Private Sub PMIC_Label_Print()

        Try
            Dim swFile As StreamWriter =
                New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))
            swFile.WriteLine("^XZ~JA^XZ")
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15^PR2,2,2") '진하기
            swFile.WriteLine("^FO0000,0000^GB684,155,2,B,0^FS")
            swFile.WriteLine("^FO0000,0055^GB684,000,2,B,0^FS")
            swFile.WriteLine("^FO0000,0090^GB684,000,2,B,0^FS")
            swFile.WriteLine("^FO0171,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0342,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0513,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0045,0015^A0,40,30^FDDetach^FS")
            swFile.WriteLine("^FO0007,0070^A0,15,15^FDDate                  /^FS")
            swFile.WriteLine("^FO0007,0100^A0,15,15^FDName^FS")
            swFile.WriteLine("^FO0195,0015^A0,40,30^FDSolder Dip^FS")
            swFile.WriteLine("^FO0178,0070^A0,15,15^FDDate                  /^FS")
            swFile.WriteLine("^FO0178,0100^A0,15,15^FDName^FS")
            swFile.WriteLine("^FO0385,0015^A0,40,30^FDAttach^FS")
            swFile.WriteLine("^FO0349,0070^A0,15,15^FDDate                  /^FS")
            swFile.WriteLine("^FO0349,0100^A0,15,15^FDName^FS")
            swFile.WriteLine("^FO0550,0015^A0,40,30^FDAttach 2^FS")
            swFile.WriteLine("^FO0520,0070^A0,15,15^FDDate                  /^FS")
            swFile.WriteLine("^FO0520,0100^A0,15,15^FDName^FS")
            swFile.WriteLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", "LPT" & TB_Port.Text)
            ElseIf cb_Cable.Text = "COM" Then
                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While

                sr.Close()
                ComPort.Close()
            ElseIf cb_Cable.Text = "USB" Then
                Dim p As New PrintRaw
                Dim s As New StringBuilder
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    s.AppendLine(str)
                End While

                sr.Close()
                p.Print(s, cb_PrinterList.Text)
            End If
            File.Delete(Application.StartupPath & "\print.txt")
        Catch ex As Exception
            File.Delete(Application.StartupPath & "\print.txt")
            MsgBox(ex.Message, MsgBoxStyle.Critical, "각종 라벨 발행")
        End Try

    End Sub

    Private Sub TestFail_Label_Print()

        If cb_Cable.Text = "USB" Then
            Dim p As New PrintRaw
            Dim s As New StringBuilder
            s.AppendLine("^XZ~JA^XZ")
            s.AppendLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            s.AppendLine("^MD15^PR2,2,2") '진하기
            s.AppendLine("^FO0065,0045^A0,80,70^FDSRB CRB Test Fail^FS")
            s.AppendLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            p.Print(s, cb_PrinterList.Text)
        Else
            Dim swFile As StreamWriter
            swFile = New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))
            'swFile.WriteLine("^XZ~JA^XZ")
            'swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            'swFile.WriteLine("^MD15^PR2,2,2") '진하기
            'swFile.WriteLine("^FO0000,0000^GB464,108,2,B,0^FS")
            'swFile.WriteLine("^FO0000,0028^GB464,000,2,B,0^FS")
            'swFile.WriteLine("^FO0000,0058^GB464,000,2,B,0^FS")
            'swFile.WriteLine("^FO0116,0000^GB000,108,2,B,0^FS")
            'swFile.WriteLine("^FO0232,0000^GB000,108,2,B,0^FS")
            'swFile.WriteLine("^FO0348,0000^GB000,108,2,B,0^FS")
            'swFile.WriteLine("^FO0030,0005^A0,25,20^FDDetach^FS")
            'swFile.WriteLine("^FO0007,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0007,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^FO0132,0005^A0,25,20^FDSolder Dip^FS")
            'swFile.WriteLine("^FO0123,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0123,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^FO0266,0005^A0,25,20^FDAttach^FS")
            'swFile.WriteLine("^FO0239,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0239,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^FO0378,0005^A0,25,20^FDScope^FS")
            'swFile.WriteLine("^FO0355,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0355,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", cb_Cable.Text & TB_Port.Text)
            Else

                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While
                sr.Close()
                ComPort.Close()
            End If
        End If
        File.Delete(Application.StartupPath & "\print.txt")

    End Sub

    Private Sub MapOut_Label_Print()

        If cb_Cable.Text = "USB" Then
            Dim p As New PrintRaw
            Dim s As New StringBuilder
            s.AppendLine("^XZ~JA^XZ")
            s.AppendLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            s.AppendLine("^MD15^PR2,2,2") '진하기
            s.AppendLine("^FO0065,0045^A0,80,70^FDMap Out^FS")
            s.AppendLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            p.Print(s, cb_PrinterList.Text)
        Else
            Dim swFile As StreamWriter
            swFile = New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))
            'swFile.WriteLine("^XZ~JA^XZ")
            'swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            'swFile.WriteLine("^MD15^PR2,2,2") '진하기
            'swFile.WriteLine("^FO0000,0000^GB464,108,2,B,0^FS")
            'swFile.WriteLine("^FO0000,0028^GB464,000,2,B,0^FS")
            'swFile.WriteLine("^FO0000,0058^GB464,000,2,B,0^FS")
            'swFile.WriteLine("^FO0116,0000^GB000,108,2,B,0^FS")
            'swFile.WriteLine("^FO0232,0000^GB000,108,2,B,0^FS")
            'swFile.WriteLine("^FO0348,0000^GB000,108,2,B,0^FS")
            'swFile.WriteLine("^FO0030,0005^A0,25,20^FDDetach^FS")
            'swFile.WriteLine("^FO0007,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0007,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^FO0132,0005^A0,25,20^FDSolder Dip^FS")
            'swFile.WriteLine("^FO0123,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0123,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^FO0266,0005^A0,25,20^FDAttach^FS")
            'swFile.WriteLine("^FO0239,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0239,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^FO0378,0005^A0,25,20^FDScope^FS")
            'swFile.WriteLine("^FO0355,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0355,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", cb_Cable.Text & TB_Port.Text)
            Else

                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While
                sr.Close()
                ComPort.Close()
            End If
        End If
        File.Delete(Application.StartupPath & "\print.txt")

    End Sub

    Private Sub PMIC_Reg_FAIL_Label_Print()

        If cb_Cable.Text = "USB" Then
            Dim p As New PrintRaw
            Dim s As New StringBuilder
            s.AppendLine("^XZ~JA^XZ")
            s.AppendLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            s.AppendLine("^MD15^PR2,2,2") '진하기
            s.AppendLine("^FO0065,0045^A0,80,70^FDPMIC Reg FAIL^FS")
            s.AppendLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            p.Print(s, cb_PrinterList.Text)
        Else
            Dim swFile As StreamWriter
            swFile = New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))
            'swFile.WriteLine("^XZ~JA^XZ")
            'swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            'swFile.WriteLine("^MD15^PR2,2,2") '진하기
            'swFile.WriteLine("^FO0000,0000^GB464,108,2,B,0^FS")
            'swFile.WriteLine("^FO0000,0028^GB464,000,2,B,0^FS")
            'swFile.WriteLine("^FO0000,0058^GB464,000,2,B,0^FS")
            'swFile.WriteLine("^FO0116,0000^GB000,108,2,B,0^FS")
            'swFile.WriteLine("^FO0232,0000^GB000,108,2,B,0^FS")
            'swFile.WriteLine("^FO0348,0000^GB000,108,2,B,0^FS")
            'swFile.WriteLine("^FO0030,0005^A0,25,20^FDDetach^FS")
            'swFile.WriteLine("^FO0007,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0007,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^FO0132,0005^A0,25,20^FDSolder Dip^FS")
            'swFile.WriteLine("^FO0123,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0123,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^FO0266,0005^A0,25,20^FDAttach^FS")
            'swFile.WriteLine("^FO0239,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0239,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^FO0378,0005^A0,25,20^FDScope^FS")
            'swFile.WriteLine("^FO0355,0036^A0,15,10^FDDate              /^FS")
            'swFile.WriteLine("^FO0355,0062^A0,15,10^FDName^FS")
            'swFile.WriteLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", cb_Cable.Text & TB_Port.Text)
            Else

                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While
                sr.Close()
                ComPort.Close()
            End If
        End If
        File.Delete(Application.StartupPath & "\print.txt")

    End Sub

    Private Sub None_Label_Print()

        Try
            Dim swFile As StreamWriter =
                New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))

            swFile.WriteLine("^XZ~JA^XZ")
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15^PR2,2,2") '진하기
            swFile.WriteLine("^FO0065,0045^A0,80,70^FDN/V^FS")
            swFile.WriteLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", "LPT" & TB_Port.Text)
            ElseIf cb_Cable.Text = "COM" Then
                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While

                sr.Close()
                ComPort.Close()
            ElseIf cb_Cable.Text = "USB" Then
                Dim p As New PrintRaw
                Dim s As New StringBuilder
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    s.AppendLine(str)
                End While

                sr.Close()
                p.Print(s, cb_PrinterList.Text)
            End If
            File.Delete(Application.StartupPath & "\print.txt")
        Catch ex As Exception
            File.Delete(Application.StartupPath & "\print.txt")
            MsgBox(ex.Message, MsgBoxStyle.Critical, "각종 라벨 발행")
        End Try

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        RadioButton5.Checked = True

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

        RadioButton6.Checked = True

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click

        RadioButton7.Checked = True

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)

        RadioButton8.Checked = True

    End Sub

    Private Sub Form_CLose_Click(sender As Object, e As EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click

        RadioButton9.Checked = True

    End Sub

    Private Sub Process1_Label_Print()

        Try
            Dim swFile As StreamWriter =
                New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))

            swFile.WriteLine("^XZ~JA^XZ")
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15") '진하기
            swFile.WriteLine("^FO0000,0000^GB684,155,2,B,0^FS")
            swFile.WriteLine("^FO010,0010^A0,30,30^FD1M^FS")
            swFile.WriteLine("^FO022,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO067,0010^A0,30,30^FD2M^FS")
            swFile.WriteLine("^FO079,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO124,0010^A0,30,30^FD3M^FS")
            swFile.WriteLine("^FO136,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO181,0010^A0,30,30^FD4M^FS")
            swFile.WriteLine("^FO193,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO238,0010^A0,30,30^FD5M^FS")
            swFile.WriteLine("^FO250,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO295,0010^A0,30,30^FD6M^FS")
            swFile.WriteLine("^FO307,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO352,0010^A0,30,30^FD7M^FS")
            swFile.WriteLine("^FO364,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO409,0010^A0,30,30^FD8M^FS")
            swFile.WriteLine("^FO421,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO466,0010^A0,30,30^FD9M^FS")
            swFile.WriteLine("^FO478,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO515,0010^A0,30,30^FD10M^FS")
            swFile.WriteLine("^FO535,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO572,0010^A0,30,30^FD11M^FS")
            swFile.WriteLine("^FO592,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO629,0010^A0,30,30^FD12M^FS")
            swFile.WriteLine("^FO649,0055^A0,30,30^FB3,10,5,L,0^FD123^FS")
            swFile.WriteLine("^FO0000,0045^GB684,000,2,B,0^FS")
            swFile.WriteLine("^FO0057,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0114,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0171,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0228,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0285,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0342,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0399,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0456,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0513,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0570,0000^GB000,155,2,B,0^FS")
            swFile.WriteLine("^FO0627,0000^GB000,155,2,B,0^FS")

            'swFile.WriteLine("^FO000,0070^A0,35,27^FDPMIC : [  ] Detach > [  ] Auto Line > [  ] Attach^FS")
            'swFile.WriteLine("^FO000,0120^A0,35,27^FD        > [  ] Reflow > [  ] AOI^FS")
            'swFile.WriteLine("^FO000,0060^A0,40,30^FDRCD  : Detach > Auto Line > Attach > Reflow > AOI^FS")
            swFile.WriteLine("^PQ" & nud_PrintQty.Value & "^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", "LPT" & TB_Port.Text)
            ElseIf cb_Cable.Text = "COM" Then
                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While

                sr.Close()
                ComPort.Close()
            ElseIf cb_Cable.Text = "USB" Then
                Dim p As New PrintRaw
                Dim s As New StringBuilder
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    s.AppendLine(str)
                End While

                sr.Close()
                p.Print(s, cb_PrinterList.Text)
            End If
            File.Delete(Application.StartupPath & "\print.txt")
        Catch ex As Exception
            File.Delete(Application.StartupPath & "\print.txt")
            MsgBox(ex.Message, MsgBoxStyle.Critical, "각종 라벨 발행")
        End Try

    End Sub

    Private Sub Process2_Label_Print()

        Try
            Dim swFile As StreamWriter =
                New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))

            swFile.WriteLine("^XZ~JA^XZ")
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15") '진하기
            swFile.WriteLine("^FO000,0000^A0,60,40^FD* Repair Process Check *^FS")
            swFile.WriteLine("^FO000,0070^A0,35,27^FDPMIC : [  ] Detach > [  ] Auto Line > [  ] Attach^FS")
            swFile.WriteLine("^FO000,0120^A0,35,27^FDRCD  : [  ] Detach > [  ] Auto Line > [  ] Reflow > [  ] AOI^FS")
            swFile.WriteLine("^PQ1^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", "LPT" & TB_Port.Text)
            ElseIf cb_Cable.Text = "COM" Then
                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While

                sr.Close()
                ComPort.Close()
            ElseIf cb_Cable.Text = "USB" Then
                Dim p As New PrintRaw
                Dim s As New StringBuilder
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    s.AppendLine(str)
                End While

                sr.Close()
                p.Print(s, cb_PrinterList.Text)
            End If
            File.Delete(Application.StartupPath & "\print.txt")
        Catch ex As Exception
            File.Delete(Application.StartupPath & "\print.txt")
            MsgBox(ex.Message, MsgBoxStyle.Critical, "각종 라벨 발행")
        End Try

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click

        RadioButton10.Checked = True

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

        RadioButton8.Checked = True

    End Sub
End Class
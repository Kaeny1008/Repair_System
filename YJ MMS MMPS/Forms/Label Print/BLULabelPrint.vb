Imports System.IO
Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient

Public Class BLULabelPrint

    Dim form_Msgbox_String As String = "Label 발행"
    Public modelGubun As String = String.Empty

    Private WithEvents ComPort As New Ports.SerialPort

    Private Sub BLULabelPrint_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GRIDSetting()
        PrintSetting()

        Panel1.Location = New Point(457, 167)

        DTP_StartDate.Value = Format(Now, "yyyy-MM-01 00:00:00")

    End Sub

    Private Sub GRIDSetting()

        With GRID_History
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .Rows(0).Height = 30
            .Rows.DefaultSize = 20
            .Cols.Count = 6
            .Cols.Fixed = 1
            .Rows.Count = 2
            .Rows.Fixed = 1
            GRID_History(0, 0) = "No"
            GRID_History(0, 1) = "발행일시"
            GRID_History(0, 2) = "구분"
            GRID_History(0, 3) = "Model Code"
            GRID_History(0, 4) = "Model Name"
            GRID_History(0, 5) = "비고"
            .AutoClipboard = True
            .ShowCursor = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            .Cols(5).StyleNew.TextAlign = TextAlignEnum.LeftCenter '비고란은 좌측 정렬
            .AutoSizeCols()
        End With

    End Sub

    Private Sub PrintSetting()

        CB_Cable.Text = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "COM/LPT", "COM")
        TB_Port.Text = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "Port", 1)
        TB_LEFT_Loc.Text = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "Width", 0)
        TB_TOP_Loc.Text = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "Height", 0)
        TB_BaudRate.Text = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "BaudRate", 9600)
        TB_DataBits.Text = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "DataBits", 8)
        TB_Parity.Text = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "Parity", 0)
        TB_StopBits.Text = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "StopBits", 1)
        TB_PrintQTY.Text = registryEdit.ReadRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "PrintQty", 1)

    End Sub

    Private Sub CB_Cable_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Cable.SelectedIndexChanged

        registryEdit.WriteRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "COM/LPT", CB_Cable.Text)

    End Sub

    Private Sub TextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Port.TextChanged,
                                                                                                        TB_LEFT_Loc.TextChanged,
                                                                                                        TB_TOP_Loc.TextChanged,
                                                                                                        TB_BaudRate.TextChanged,
                                                                                                        TB_DataBits.TextChanged,
                                                                                                        TB_Parity.TextChanged,
                                                                                                        TB_StopBits.TextChanged,
                                                                                                        TB_PrintQTY.TextChanged

        Dim abcd As TextBox = CType(sender, TextBox)

        Select Case abcd.Name
            Case TB_Port.Name
                registryEdit.WriteRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "Port", TB_Port.Text)
            Case TB_LEFT_Loc.Name
                registryEdit.WriteRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "Width", TB_LEFT_Loc.Text)
            Case TB_TOP_Loc.Name
                registryEdit.WriteRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "Height", TB_TOP_Loc.Text)
            Case TB_BaudRate.Name
                registryEdit.WriteRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "BaudRate", TB_BaudRate.Text)
            Case TB_DataBits.Name
                registryEdit.WriteRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "DataBits", TB_DataBits.Text)
            Case TB_Parity.Name
                registryEdit.WriteRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "Parity", TB_Parity.Text)
            Case TB_StopBits.Name
                registryEdit.WriteRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "StopBits", TB_StopBits.Text)
            Case TB_PrintQTY.Name
                registryEdit.WriteRegKey("Software\Yujin\MMS_MMPS\Label\Printer", "PrintQty", TB_PrintQTY.Text)
        End Select

    End Sub

    Private Sub TextBox_Loc_KeyDown(sender As Object, e As KeyEventArgs) Handles TB_Port.KeyDown,
                                                                                 TB_LEFT_Loc.KeyDown,
                                                                                 TB_TOP_Loc.KeyDown,
                                                                                 TB_BaudRate.KeyDown,
                                                                                 TB_DataBits.KeyDown,
                                                                                 TB_Parity.KeyDown,
                                                                                 TB_StopBits.KeyDown,
                                                                                 TB_PrintQTY.KeyDown

        Select Case e.KeyCode
            Case 48 To 57, 96 To 105 'number
            Case 13 'enter
            Case 8 'backspace
            Case Else
                e.SuppressKeyPress = True
        End Select

    End Sub

    Private Sub TextBox_Leave(sender As Object, e As EventArgs) Handles TB_Port.Leave,
                                                                        TB_LEFT_Loc.Leave,
                                                                        TB_TOP_Loc.Leave,
                                                                        TB_BaudRate.Leave,
                                                                        TB_DataBits.Leave,
                                                                        TB_Parity.Leave,
                                                                        TB_StopBits.Leave,
                                                                        TB_PrintQTY.Leave

        Dim abcd As TextBox = CType(sender, TextBox)

        If abcd.Text = String.Empty Then
            abcd.Text = 0
        End If

    End Sub

    Private Sub BTN_ModelSearch_Click(sender As Object, e As EventArgs) Handles BTN_ModelSearch.Click

        If Not ModelSearch.Visible Then ModelSearch.Show()
        ModelSearch.Focus()

    End Sub

    Private Sub BTN_PrinterSetting_Click(sender As Object, e As EventArgs) Handles BTN_PrinterSetting.Click

        If BTN_PrinterSetting.Text = "프린터 설정 보기" Then
            BTN_PrinterSetting.Text = "프린터 설정 닫기"
            Panel1.Visible = True
        Else
            BTN_PrinterSetting.Text = "프린터 설정 보기"
            Panel1.Visible = False
        End If

    End Sub

    Private Sub TB_ModelName_KeyDown(sender As Object, e As KeyEventArgs) Handles TB_ModelName.KeyDown

        If e.KeyCode = 13 And Not TB_ModelName.Text = String.Empty Then
            ModelSearch.findText = TB_ModelName.Text
            If Not ModelSearch.Visible Then ModelSearch.Show()
            ModelSearch.Focus()
        End If

    End Sub

    Private Sub BTN_LabelPrint_Click(sender As Object, e As EventArgs) Handles BTN_LabelPrint.Click

        If TB_ModelCode.Text = String.Empty Then
            MsgBox("모델을 선택하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If MsgBox("Label을 인쇄 하시겠습니까?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        If Label_Print_Run() = False Then
            MsgBox("Label 인쇄 실패", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand

        sqlTran = DBConnect1.BeginTransaction

        Try

            Dim writeDate As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            Dim strSQL As String = "Insert into TB_PRINT_HISTORY(DATE_TIME, MODEL_GUBUN, MODEL_CODE, MODEL_NAME, HISTORY_NOTE) values"
            strSQL += "('" & writeDate & "'"
            strSQL += ",'" & modelGubun & "'"
            strSQL += ",'" & TB_ModelCode.Text & "'"
            strSQL += ",'" & TB_ModelName.Text & "'"
            strSQL += ",'" & TB_HistoryNote.Text & "')"

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, DBConnect1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        MsgBox("저장완료.", MsgBoxStyle.Information, form_Msgbox_String)

        modelGubun = String.Empty
        TB_ModelCode.Text = String.Empty
        TB_ModelName.Text = String.Empty
        TB_HistoryNote.Text = String.Empty

        BTN_Search_Click(Nothing, Nothing)

    End Sub

    Private Function Label_Print_Run() As Boolean

        Label_Print_Run = False

        ComPort.PortName = CB_Cable.Text & TB_Port.Text
        ComPort.BaudRate = TB_BaudRate.Text
        ComPort.Parity = TB_Parity.Text
        ComPort.DataBits = TB_DataBits.Text
        ComPort.StopBits = TB_StopBits.Text
        ComPort.Encoding = System.Text.Encoding.Default

        Try
            Dim swFile As StreamWriter
            swFile = New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))
            swFile.WriteLine("^XZ~JA^XZ")
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15") '진하기
            swFile.WriteLine("^FO200,010^BY2,3,10^BCN,050,N,N,N^FD" & "BLU-" & TB_ModelCode.Text & "^FS") '1D Barcode : CODE 128
            swFile.WriteLine("^FO200,070^A0,32,25^FD" & TB_ModelName.Text & "^FS")
            swFile.WriteLine("^PQ" & TB_PrintQTY.Text & "^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If CB_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", CB_Cable.Text & TB_Port.Text)
            Else
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
            File.Delete(Application.StartupPath & "\print.txt")
            Label_Print_Run = True
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Information, "Label Print")
        End Try

        Return Label_Print_Run

    End Function

    Private Sub BTN_Search_Click(sender As Object, e As EventArgs) Handles BTN_Search.Click

        GRID_History.Redraw = False
        GRID_History.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call USP_HISTORY(5"
        strSQL += ",'" & Format(DTP_StartDate.Value, "yyyy-MM-dd 00:00:00") & "'"
        strSQL += ",'" & Format(DTP_EndDate.Value, "yyyy-MM-dd 23:59:59") & "'"
        strSQL += ",null"
        strSQL += ",'" & TB_ModelName.Text & "'"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null"
        strSQL += ",null)"

        Dim sqlCmd As New MySqlCommand(strSQL, DBConnect1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim insert As String = String.Empty
            insert = GRID_History.Rows.Count &
                vbTab & Format(sqlDR("DATE_TIME"), "yyyy-MM-dd HH:mm:ss") &
                vbTab & sqlDR("MODEL_GUBUN") &
                vbTab & sqlDR("MODEL_CODE") &
                vbTab & sqlDR("MODEL_NAME") &
                vbTab & sqlDR("HISTORY_NOTE")
            GRID_History.AddItem(insert)
        Loop
        sqlDR.Close()

        DBClose()

        GRID_History.AutoSizeCols()
        GRID_History.AutoSizeRows()
        GRID_History.Redraw = True

    End Sub

    Private Sub BLULabelPrint_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        Mainform.form_name = Me

    End Sub
End Class
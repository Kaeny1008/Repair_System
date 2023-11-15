Imports System.Xml
Imports C1.Win.C1FlexGrid

Public Class CadData
    Private Sub BTN_CadData_Click(sender As Object, e As EventArgs) Handles BTN_CadData.Click

        'xml파일을 불러와서 보여준다.

        'Dim openFileDialog1 As New OpenFileDialog

        'openFileDialog1.FileName = String.Empty

        'openFileDialog1.Filter = "XML Files (*.xml)|*.xml"
        'openFileDialog1.ShowDialog()

        'If openFileDialog1.FileName = String.Empty Then Exit Sub

        'Dim doc = New XmlDocument
        'doc.Load(openFileDialog1.FileName)

        'Dim root As XmlNode = doc.DocumentElement

        'Dim devices As XmlNodeList = root.SelectNodes("PCBDATA/MACHINE/DEVICES/FEEDERBASE")
        'Dim machine_no As Integer = 1
        'If IsNumeric(GRID_LIST_DEVICE(GRID_LIST_DEVICE.Rows.Count - 1, 2)) Then
        '    machine_no = CInt(GRID_LIST_DEVICE(GRID_LIST_DEVICE.Rows.Count - 1, 2)) + 1
        'Else
        '    If GRID_LIST_DEVICE(GRID_LIST_DEVICE.Rows.Count - 1, 2) = String.Empty Then GRID_LIST_DEVICE.Rows.Count = 1
        'End If
        'Dim now_base_no As Integer = 1

        'For Each base_no As XmlNode In devices
        '    Dim feeder_no As XmlNodeList = base_no.SelectNodes("TAPEFEEDER")
        '    If Not now_base_no = base_no.Attributes("ID").Value Then
        '        machine_no += 1
        '    End If
        '    For Each feeder_info As XmlNode In feeder_no
        '        'Console.WriteLine(machine_no & ",   " & _
        '        '                  feeder_info.SelectSingleNode("INSTALL").Attributes("SLOT").Value & ",   " & _
        '        '                  feeder_info.SelectSingleNode("TAPE").Attributes("PARTNAME").Value)
        '        Dim insert_string As String = "N" & vbTab &
        '              NEW_Code() & vbTab &
        '              machine_no & vbTab &
        '              feeder_info.SelectSingleNode("INSTALL").Attributes("SLOT").Value & vbTab &
        '              String.Empty & vbTab &
        '              feeder_info.SelectSingleNode("TAPE").Attributes("PARTNAME").Value & vbTab &
        '              String.Empty & vbTab &
        '              String.Empty & vbTab &
        '              String.Empty & vbTab &
        '              String.Empty
        '        GRID_LIST_DEVICE.AddItem(insert_string)

        '    Next
        'Next

        'MsgBox("Device Data를 불러 왔습니다.", MsgBoxStyle.Information, strMSG)

        GridSetting("ref")

        Grid_CADData.Rows.Count = 1

        Dim openFileDialog1 As New OpenFileDialog

        openFileDialog1.FileName = String.Empty

        openFileDialog1.Filter = "XML Files (*.xml)|*.xml"
        openFileDialog1.ShowDialog()

        If openFileDialog1.FileName = String.Empty Then Exit Sub

        Dim doc = New XmlDocument
        doc.Load(openFileDialog1.FileName)

        Dim root As XmlNode = doc.DocumentElement

        Dim listCadData As XmlNodeList = root.SelectNodes("PCBDATA/CADDATA")

        For Each base_no As XmlNode In listCadData
            Dim feeder_no As XmlNodeList = base_no.SelectNodes("POINT")
            For Each abcd As XmlNode In feeder_no
                Grid_CADData.AddItem(Grid_CADData.Rows.Count & vbTab &
                                     abcd.Attributes("REF").Value & vbTab &
                                     abcd.Attributes("PART").Value)
            Next
        Next

        Grid_CADData.AutoSizeRows()
        Grid_CADData.AutoSizeCols()

        MsgBox("Cad Data를 불러 왔습니다.", MsgBoxStyle.Information, "SAMSUNG Cad Data Load")

    End Sub

    Private Sub CadData_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GridSetting("ref")

    End Sub

    Private Sub GridSetting(ByVal section As String)

        If section = "ref" Then
            With Grid_CADData
                .AllowEditing = True
                .AllowFiltering = True
                .AutoClipboard = True
                .AllowSorting = AllowSortingEnum.None
                .Rows(0).Height = 30
                .Rows.DefaultSize = 20
                .Cols.Count = 3
                .Cols.Fixed = 1
                .Rows.Count = 1
                .Rows.Fixed = 1
                Grid_CADData(0, 0) = "No"
                Grid_CADData(0, 1) = "Ref"
                Grid_CADData(0, 2) = "Part No."
                .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
                .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
                .AutoSizeCols()
                .ExtendLastCol = True
            End With
        ElseIf section = "feeder" Then
            With Grid_CADData
                .AllowEditing = True
                .AllowFiltering = True
                .AutoClipboard = True
                .AllowSorting = AllowSortingEnum.None
                .Rows(0).Height = 30
                .Rows.DefaultSize = 20
                .Cols.Count = 4
                .Cols.Fixed = 1
                .Rows.Count = 1
                .Rows.Fixed = 1
                Grid_CADData(0, 0) = "No"
                Grid_CADData(0, 1) = "Feeder Type"
                Grid_CADData(0, 2) = "Slot No."
                Grid_CADData(0, 3) = "Part No."
                .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
                .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
                .AutoSizeCols()
                .ExtendLastCol = True
            End With
        End If

    End Sub

    Private Sub BTN_FeederLoad_Click(sender As Object, e As EventArgs) Handles BTN_FeederLoad.Click

        GridSetting("feeder")

        Grid_CADData.Rows.Count = 1

        Dim openFileDialog1 As New OpenFileDialog

        openFileDialog1.FileName = String.Empty

        openFileDialog1.Filter = "XML Files (*.xml)|*.xml"
        openFileDialog1.ShowDialog()

        If openFileDialog1.FileName = String.Empty Then Exit Sub

        Dim doc = New XmlDocument
        doc.Load(openFileDialog1.FileName)

        Dim root As XmlNode = doc.DocumentElement

        Dim devices As XmlNodeList = root.SelectNodes("PCBDATA/MACHINE/DEVICES/FEEDERBASE")

        For Each base_no As XmlNode In devices
            Dim feeder_no As XmlNodeList = base_no.SelectNodes("TAPEFEEDER")
            For Each feeder_info As XmlNode In feeder_no
                'Dim insert_string As String = "N" & vbTab &
                '      NEW_Code() & vbTab &
                '      machine_no & vbTab &
                '      feeder_info.SelectSingleNode("INSTALL").Attributes("SLOT").Value & vbTab &
                '      String.Empty & vbTab &
                '      feeder_info.SelectSingleNode("TAPE").Attributes("PARTNAME").Value & vbTab &
                '      String.Empty & vbTab &
                '      String.Empty & vbTab &
                '      String.Empty & vbTab &
                '      String.Empty
                Dim insert_string As String = Grid_CADData.Rows.Count & vbTab &
                    feeder_info.Attributes("TYPE").Value & vbTab &
                    feeder_info.SelectSingleNode("INSTALL").Attributes("SLOT").Value & vbTab &
                    feeder_info.SelectSingleNode("TAPE").Attributes("PARTNAME").Value
                Grid_CADData.AddItem(insert_string)
            Next
        Next

        Grid_CADData.AutoSizeRows()
        Grid_CADData.AutoSizeCols()

        MsgBox("Cad Data를 불러 왔습니다.", MsgBoxStyle.Information, "SAMSUNG Cad Data Load")

    End Sub
End Class
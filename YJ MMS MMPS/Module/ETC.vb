Module ETC

    Public Sub MMPS_DB_Registry()

        If registryEdit.ReadRegKey("Software\ODBC\ODBC.INI\MMPS_DB", "DBQ", String.Empty) = String.Empty Then
            registryEdit.WriteRegKey("Software\ODBC\ODBC.INI\MMPS_DB", "DBQ", ".\TempDB\TempDB_MMPS.mdb")
            registryEdit.WriteRegKey("Software\ODBC\ODBC.INI\MMPS_DB", "Driver", "C:\windows\system32\odbcjt32.dll")
            registryEdit.WriteRegKey("Software\ODBC\ODBC.INI\MMPS_DB", "DriverId", 25)
            registryEdit.WriteRegKey("Software\ODBC\ODBC.INI\MMPS_DB", "FIL", "MS Access;")
            registryEdit.WriteRegKey("Software\ODBC\ODBC.INI\MMPS_DB", "SafeTransactions", 0)
            registryEdit.WriteRegKey("Software\ODBC\ODBC.INI\MMPS_DB", "UID", String.Empty)
        End If

    End Sub

    Public Function Barcode_Split(ByVal barcode_string As String,
                                  ByVal maker As String)

        Try
            If maker.ToUpper = "MOLEX" Then
                '1P로 시작은 Part No.
                If barcode_string.ToUpper.Substring(0, 2) = "1P" Then
                    barcode_string = "P:" & barcode_string.Substring(2, barcode_string.Length - 2)
                ElseIf barcode_string.ToUpper.Substring(0, 1) = "Q" Then
                    barcode_string = "Q:" & barcode_string.Substring(1, barcode_string.Length - 1)
                ElseIf barcode_string.ToUpper.Substring(0, 2) = "1T" Then
                    barcode_string = "L:" & barcode_string.Substring(4, barcode_string.Length - 4)
                End If
            ElseIf maker.ToUpper = "ELMOS" Then
                If barcode_string.ToUpper.Substring(0, 1) = "P" Then
                    barcode_string = "P:" & barcode_string.Substring(1, barcode_string.Length - 2)
                ElseIf barcode_string.ToUpper.Substring(0, 1) = "Q" Then
                    barcode_string = "Q:" & barcode_string.Substring(1, barcode_string.Length - 2)
                ElseIf barcode_string.ToUpper.Substring(0, 1) = "D" Then
                    barcode_string = barcode_string.Substring(1, barcode_string.Length - 2)
                ElseIf barcode_string.ToUpper.Substring(0, 1) = "H" Then
                    barcode_string = barcode_string.Substring(1, barcode_string.Length - 2)
                ElseIf barcode_string.ToUpper.Substring(0, 1) = "S" Then
                    barcode_string = "L:" & barcode_string.Substring(1, barcode_string.Length - 2)
                End If
            ElseIf maker.ToUpper = "MURATA" Then
                If barcode_string.ToUpper.Substring(0, 1) = "P" Then
                    barcode_string = "P:" & barcode_string.Substring(1, barcode_string.Length - 1)
                ElseIf barcode_string.ToUpper.Substring(0, 2) = "1P" Then
                    barcode_string = "P:" & barcode_string.Substring(2, barcode_string.Length - 2)
                ElseIf barcode_string.ToUpper.Substring(0, 1) = "Q" Then
                    barcode_string = "Q:" & barcode_string.Substring(1, barcode_string.Length - 1)
                ElseIf barcode_string.ToUpper.Substring(0, 2) = "1T" Then
                    barcode_string = "L:" & barcode_string.Substring(2, barcode_string.Length - 2)
                End If
            ElseIf maker.ToUpper = "NEXPERIA" Then
                If barcode_string.ToUpper.Substring(0, 2) = "1P" Then
                    barcode_string = "P:" & barcode_string.Substring(2, barcode_string.Length - 2)
                ElseIf barcode_string.ToUpper.Substring(0, 1) = "Q" Then
                    barcode_string = "Q:" & barcode_string.Substring(1, barcode_string.Length - 1)
                ElseIf barcode_string.ToUpper.Substring(0, 2) = "1T" Then
                    barcode_string = "L:" & barcode_string.Substring(2, barcode_string.Length - 2)
                End If
            ElseIf maker.ToUpper = "YAGEO" Then
                If barcode_string.ToUpper.Substring(0, 1) = "P" Then
                    barcode_string = "P:" & barcode_string.Substring(1, barcode_string.Length - 1)
                ElseIf barcode_string.ToUpper.Substring(0, 2) = "1P" Then
                    barcode_string = "P:" & barcode_string.Substring(2, barcode_string.Length - 2)
                ElseIf barcode_string.ToUpper.Substring(0, 1) = "Q" Then
                    barcode_string = "Q:" & barcode_string.Substring(1, barcode_string.Length - 1)
                ElseIf barcode_string.ToUpper.Substring(0, 1) = "K" Then
                    barcode_string = "L:" & barcode_string.Substring(1, barcode_string.Length - 1)
                End If
            ElseIf maker.ToUpper = "WALSIN" Then
                Dim replace_string As String = String.Empty
                If barcode_string.Length = 31 Then 'Part, Lot
                    replace_string = barcode_string.Substring(0, 12)
                    replace_string += "!@" & barcode_string.Substring(16, barcode_string.Length - 18)
                    replace_string += "!@" & barcode_string.Substring(29, barcode_string.Length - 29) * 1000
                    barcode_string = replace_string
                End If
            ElseIf maker.ToUpper = "OSRAM" Then
                If InStr(barcode_string, "@") > 0 Then
                    Dim split_text() As String = Split(barcode_string, "@")
                    Dim part_no As String = String.Empty
                    Dim lot_no As String = String.Empty
                    Dim qty As String = String.Empty
                    For i = 0 To UBound(split_text)
                        If split_text(i).Substring(0, 2) = "1P" And split_text(i).Length > 2 Then
                            part_no = Replace(split_text(i), "1P", String.Empty)
                        ElseIf split_text(i).Substring(0, 2) = "1T" Then
                            lot_no = Replace(split_text(i), "1T", String.Empty)
                        ElseIf split_text(i).Substring(0, 1) = "Q" Then
                            qty = Replace(split_text(i), "Q", String.Empty)
                        End If
                    Next
                    barcode_string = part_no & "!@" & lot_no & "!@" & qty
                End If
            ElseIf maker.ToUpper = "VISHAY" Then
                If Len(barcode_string) > 99 Then
                    Dim part_no As String = Trim(barcode_string.Substring(0, 18))
                    Dim lot_no As String = Trim(barcode_string.Substring(89, 10))
                    Dim qty As Integer = CInt(Trim(barcode_string.Substring(99, 13)))
                    barcode_string = part_no & "!@" & lot_no & "!@" & qty
                End If
            ElseIf maker.ToUpper = "SAMSUNG" Then
                If InStr(barcode_string, "/") > 0 Then
                    Dim part_no As String = barcode_string.Split("/")(1)
                    Dim lot_no As String = barcode_string.Split("/")(0)
                    Dim qty As Integer = CInt(barcode_string.Split("/")(3))
                    barcode_string = part_no & "!@" & lot_no & "!@" & qty
                End If
            ElseIf maker.ToUpper = "SUNLOAD" Then
                'RLE1907290238/SZ1608G331TF/FC5X03971/5/37/4000/CG90171/K7V0VDC
                If InStr(barcode_string, "/") > 0 Then
                    Dim part_no As String = barcode_string.Split("/")(0)
                    Dim lot_no As String = barcode_string.Split("/")(2)
                    Dim qty As Integer = CInt(barcode_string.Split("/")(7))
                    barcode_string = part_no & "!@" & lot_no & "!@" & qty
                End If
            ElseIf maker.ToUpper = "UNIOHM" Then
                '810261770109-0402J0331
                'RC11708271962R
                '0RJ1000C678!812!0402WGJ0101TCE!10610981
                If InStr(barcode_string, "!") > 0 Then
                    Dim part_no As String = barcode_string.Split("!")(2)
                    Dim lot_no As String = barcode_string.Split("!")(3)
                    barcode_string = part_no & "!@" & lot_no
                ElseIf InStr(barcode_string, "-") > 0 Then
                    Dim part_no As String = barcode_string.Split("-")(1)
                    barcode_string = part_no
                End If
            ElseIf maker.ToUpper = "YEONHO" Then
                '3711-007975DF141906180624002000
                If Len(barcode_string) = 31 Then
                    Dim part_no As String = barcode_string.Substring(0, 11)
                    Dim lot_no As String = barcode_string.Substring(11, 14)
                    Dim qty As Integer = CInt(barcode_string.Substring(25, 6))
                    barcode_string = part_no & "!@" & lot_no & "!@" & qty
                End If
            ElseIf maker.ToUpper = "ALPS" Then
                '[)>0611Z00112ZAT040TZATS040224161971000002989V1007088PSKRMAAE0101PSKRMAAE010Q100001TJ449710ZAB1K2559698821LJPK60040190244K1017Z10

                'Dim split_text() As String = Split(barcode_string, Chr(29))
                'Dim part_no As String = String.Empty
                'Dim lot_no As String = String.Empty
                'Dim qty As String = String.Empty

                'MsgBox(InStr(barcode_string, Chr(29)))

                'For i = 0 To UBound(split_text)
                '    If split_text(i).Substring(0, 2) = "1P" And split_text(i).Length > 2 Then
                '        part_no = Replace(split_text(i), "1P", String.Empty)
                '    ElseIf split_text(i).Substring(0, 2) = "1T" Then
                '        lot_no = Replace(split_text(i), "1T", String.Empty)
                '    ElseIf split_text(i).Substring(0, 1) = "Q" Then
                '        qty = Replace(split_text(i), "Q", String.Empty)
                '    End If
                'Next
                Dim loc_1P As Integer = InStr(barcode_string, "1P")
                Dim loc_Q As Integer = InStr(barcode_string, "Q")
                Dim loc_1T As Integer = InStr(barcode_string, "1T")
                Dim loc_1K As Integer = InStr(barcode_string, "1K")

                Dim part_no As String = barcode_string.Substring(loc_1P + 1, loc_Q - loc_1P - 2)
                Dim lot_no As String = barcode_string.Substring(loc_1T + 1, loc_1K - loc_1T - 2)
                Dim qty As Integer = barcode_string.Substring(loc_Q, loc_1T - loc_Q - 1)

                barcode_string = part_no & "!@" & lot_no & "!@" & qty
            End If
        Catch ex As Exception

        End Try

        Return barcode_string

    End Function

    Public Function CInt2(ByVal Value As String)

        If Trim(Value) = "" Then
            CInt2 = 0
        Else
            CInt2 = CDbl(Value)
        End If

    End Function

    Public Function RELotNoMaking(ByVal Number As String) As String

        RELotNoMaking = ""

        Select Case Number
            Case "0"
                RELotNoMaking += "0"
            Case "1"
                RELotNoMaking += "1"
            Case "2"
                RELotNoMaking += "2"
            Case "3"
                RELotNoMaking += "3"
            Case "4"
                RELotNoMaking += "4"
            Case "5"
                RELotNoMaking += "5"
            Case "6"
                RELotNoMaking += "6"
            Case "7"
                RELotNoMaking += "7"
            Case "8"
                RELotNoMaking += "8"
            Case "9"
                RELotNoMaking += "9"
            Case "A"
                RELotNoMaking += "10"
            Case "B"
                RELotNoMaking += "11"
            Case "C"
                RELotNoMaking += "12"
            Case "D"
                RELotNoMaking += "13"
            Case "E"
                RELotNoMaking += "14"
            Case "F"
                RELotNoMaking += "15"
            Case "G"
                RELotNoMaking += "16"
            Case "H"
                RELotNoMaking += "17"
            Case "J"
                RELotNoMaking += "18"
            Case "K"
                RELotNoMaking += "19"
            Case "L"
                RELotNoMaking += "20"
            Case "M"
                RELotNoMaking += "21"
            Case "N"
                RELotNoMaking += "22"
            Case "P"
                RELotNoMaking += "23"
            Case "R"
                RELotNoMaking += "24"
            Case "S"
                RELotNoMaking += "25"
            Case "T"
                RELotNoMaking += "26"
            Case "U"
                RELotNoMaking += "27"
            Case "V"
                RELotNoMaking += "28"
            Case "W"
                RELotNoMaking += "29"
            Case "X"
                RELotNoMaking += "30"
            Case "Z"
                RELotNoMaking += "31"

        End Select

        Return RELotNoMaking

    End Function

    Public Function LotNoMaking(ByVal Number As Integer) As String

        LotNoMaking = ""

        Select Case Number
            Case 0
                LotNoMaking += "0"
            Case 1
                LotNoMaking += "1"
            Case 2
                LotNoMaking += "2"
            Case 3
                LotNoMaking += "3"
            Case 4
                LotNoMaking += "4"
            Case 5
                LotNoMaking += "5"
            Case 6
                LotNoMaking += "6"
            Case 7
                LotNoMaking += "7"
            Case 8
                LotNoMaking += "8"
            Case 9
                LotNoMaking += "9"
            Case 10
                LotNoMaking += "A"
            Case 11
                LotNoMaking += "B"
            Case 12
                LotNoMaking += "C"
            Case 13
                LotNoMaking += "D"
            Case 14
                LotNoMaking += "E"
            Case 15
                LotNoMaking += "F"
            Case 16
                LotNoMaking += "G"
            Case 17
                LotNoMaking += "H"
            Case 18
                LotNoMaking += "J"
            Case 19
                LotNoMaking += "K"
            Case 20
                LotNoMaking += "L"
            Case 21
                LotNoMaking += "M"
            Case 22
                LotNoMaking += "N"
            Case 23
                LotNoMaking += "P"
            Case 24
                LotNoMaking += "R"
            Case 25
                LotNoMaking += "S"
            Case 26
                LotNoMaking += "T"
            Case 27
                LotNoMaking += "U"
            Case 28
                LotNoMaking += "V"
            Case 29
                LotNoMaking += "W"
            Case 30
                LotNoMaking += "X"
            Case 31
                LotNoMaking += "Z"

        End Select

        Return LotNoMaking

    End Function

End Module

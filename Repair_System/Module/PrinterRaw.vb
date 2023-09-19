Imports System.Runtime.InteropServices
Imports System.Text

Public Class PrintRaw

    Public Sub Print(ByVal codes As StringBuilder, ByVal printer As String)
        SendToPrinter("Printing Zebras", codes.ToString, printer)
    End Sub

    Private Structure Docinfo
        <MarshalAs(UnmanagedType.LPWStr)> Public DocumentName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public OutputFile As String
        <MarshalAs(UnmanagedType.LPWStr)> Public DataType As String
    End Structure

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=False, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function OpenPrinter(ByVal pPrinterName As String, ByRef phPrinter As IntPtr, ByVal pDefault As Integer) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=False, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function StartDocPrinter(ByVal hPrinter As IntPtr, ByVal level As Integer, ByRef pDocInfo As Docinfo) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function StartPagePrinter(ByVal hPrinter As IntPtr) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Ansi, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function WritePrinter(ByVal hPrinter As IntPtr, ByVal data As String, ByVal buf As Integer, ByRef pcWritten As Integer) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function EndPagePrinter(ByVal hPrinter As IntPtr) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function EndDocPrinter(ByVal hPrinter As IntPtr) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function ClosePrinter(ByVal hPrinter As IntPtr) As Long
    End Function

    Private Sub SendToPrinter(ByVal printerJobName As String, ByVal rawStringToSendToThePrinter As String, ByVal printerNameAsDescribedByPrintManager As String)

        Dim handleForTheOpenPrinter = New IntPtr()
        Dim documentInformation = New Docinfo()
        Dim printerBytesWritten = 0
        documentInformation.DocumentName = printerJobName
        documentInformation.DataType = vbNullString
        documentInformation.OutputFile = vbNullString
        OpenPrinter(printerNameAsDescribedByPrintManager, handleForTheOpenPrinter, 0)
        StartDocPrinter(handleForTheOpenPrinter, 1, documentInformation)
        StartPagePrinter(handleForTheOpenPrinter)
        WritePrinter(handleForTheOpenPrinter, rawStringToSendToThePrinter, rawStringToSendToThePrinter.Length, printerBytesWritten)
        EndPagePrinter(handleForTheOpenPrinter)
        EndDocPrinter(handleForTheOpenPrinter)
        ClosePrinter(handleForTheOpenPrinter)

    End Sub

End Class
Imports System.IO
Imports System.Net

Public Class frm_Main

    Dim RegistryEdit As New RegistryEdit.RegReadWrite
    Dim ftpUrl As String = "ftp://" & RegistryEdit.ReadRegKey("Software\Yujin\Repair System", "Server.IP", "192.168.0.222") & ":" & RegistryEdit.ReadRegKey("Software\Yujin\Repair System\FTP", "ftpPort", "1052")
    Dim ftpPort As Integer = RegistryEdit.ReadRegKey("Software\Yujin\Repair System\FTP", "ftpPort", "1052")
    Dim ftpID As String = "yujin_ftp"
    Dim ftpPassword As String = "Yujin_ftp!"
    Dim i As Integer
    Dim CheckFileName As String
    Dim FTPDownloading As Boolean = False

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Timer1.Interval = 1000
        Timer1.Enabled = True

        If UBound(Diagnostics.Process.GetProcessesByName("UpdateChecker")) = 0 Then
            'UpdateChecker가 켜져 있다면 종료
            Dim ucProcess() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("UpdateChecker")
            If ucProcess.Length > 0 Then
                ucProcess(0).Kill()
            End If
        End If

    End Sub

    '************************* FTP 서버내 Ver. CheckFile 다운로드 ***************************
    Private Sub CheckVerDownload(ByVal ftpURL As String, ByVal FileName As String, ByVal RunContinue As Boolean)

        Label1.Text = "Version Check File 다운로드....."
        '############################  파일다운로드 관련부분 #########################################################

        Dim localPath_download As String = Application.StartupPath & "\" '다운로드 받을 경로

        'While fileName IsNot Nothing '무조건 다운로드할때 이걸써라.(항목수 상관없이)
        Dim requestFileDelete As FtpWebRequest = Nothing
        Dim responseFileDelete As FtpWebResponse = Nothing

        Dim requestFileDownload As FtpWebRequest = Nothing
        Dim responseFileDownload As FtpWebResponse = Nothing

        '파일존재여부 확인
        'If My.Computer.FileSystem.FileExists(localPath_download & "\" & fileName) Then
        '    Try
        '        Kill(localPath_download & "\" & fileName)
        '    Catch ex As Exception
        '        MsgBox(ex.Message, MsgBoxStyle.Critical, Loginsite_Hangul)
        '    End Try
        'End If

        Try
            'Console.WriteLine(fileName)

            requestFileDownload = DirectCast(WebRequest.Create(ftpURL & FileName), FtpWebRequest)
            requestFileDownload.Credentials = New NetworkCredential(ftpID, ftpPassword)
            requestFileDownload.Method = WebRequestMethods.Ftp.DownloadFile
            requestFileDownload.UsePassive = True
            requestFileDownload.KeepAlive = False
            requestFileDownload.UseBinary = False
            responseFileDownload = DirectCast(requestFileDownload.GetResponse(), FtpWebResponse)

            Dim responseStream As Stream = responseFileDownload.GetResponseStream()
            Dim writeStream As New FileStream(localPath_download & FileName, FileMode.Create)
            Dim Length As Integer = 2048
            Dim buffer As [Byte]() = New [Byte](Length - 1) {}
            Dim bytesRead As Integer = responseStream.Read(buffer, 0, Length)

            '#####파일의 크기를 알아보기 위해...#####
            Dim requestFileSize As FtpWebRequest = Nothing
            requestFileSize = DirectCast(WebRequest.Create(ftpURL & "/" & FileName), FtpWebRequest)
            requestFileSize.Credentials = New NetworkCredential(ftpID, ftpPassword)
            requestFileSize.Method = WebRequestMethods.Ftp.GetFileSize
            '프로그래스바의 최대를 알아본다.
            'ProgressBar1.Maximum = requestFileSize.GetResponse.ContentLength
            Dim run_count As Integer = 0
            Dim total_count As Integer = requestFileSize.GetResponse.ContentLength
            ProgressBar1.Maximum = 100
            ProgressBar1.Value = 1
            requestFileSize = Nothing

            While bytesRead > 0
                writeStream.Write(buffer, 0, bytesRead)
                bytesRead = responseStream.Read(buffer, 0, Length)
                'ProgressBar1.Value += bytesRead
                run_count += bytesRead
                ProgressBar1.Value = run_count / total_count * 100
                Application.DoEvents()
            End While
            writeStream.Close()

        Catch exceptionObj As Exception
            MsgBox(exceptionObj.Message.ToString())
        Finally
            responseFileDelete = Nothing
            requestFileDelete = Nothing

            requestFileDownload = Nothing
            responseFileDownload = Nothing
        End Try

        'End While

        '############################  파일다운로드 관련부분 #########################################################
        Label1.Text = "Version Check File 다운로드 완료."

        If RunContinue = True Then
            CheckVerTextOpen()
        End If

    End Sub

    Private Sub CheckVerTextOpen()

        Label1.Text = "Version 비교중....."

        Dim ReadOk As Integer
        Dim ReadOk2 As Integer
        '새로운파일버젼 확인
        OpenFileDialog1.FileName = Application.StartupPath & "\NEWCheckVer.ini"

        Dim fs As System.IO.FileStream
        Dim sr As System.IO.StreamReader
        fs = System.IO.File.Open(OpenFileDialog1.FileName, IO.FileMode.Open) ' 파일 열기
        sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결

        Dim str As String
        Dim NewVer() As String
        Dim OldVer() As String

        'While sr.EndOfStream <> False ' 끝까지 읽기
        While sr.Peek() >= 0
            str = sr.ReadLine() ' 한줄씩 읽기
            If str = "[AUTO UPDATE]" Then ReadOk = 1 '파일읽기 시작
            If str = "[/]" Then ReadOk = 0 '파일읽기 종료

            NewVer = Split(str, "=")

            If ReadOk = 1 And NewVer.Length > 1 Then '읽시 시작이 Ok면
                OpenFileDialog1.FileName = Application.StartupPath & "\CheckVer.ini"

                Dim sr1 As System.IO.StreamReader
                Dim fs1 As System.IO.FileStream
                fs1 = System.IO.File.Open(OpenFileDialog1.FileName, IO.FileMode.Open) ' 파일 열기
                sr1 = New System.IO.StreamReader(fs1) ' 스트림리더에 연결

                Dim Download As Boolean = True

                While sr1.Peek() >= 0
                    str = sr1.ReadLine() ' 한줄씩 읽기
                    If str = "[AUTO UPDATE]" Then ReadOk2 = 1 '파일읽기 시작
                    If str = "[/]" Then ReadOk2 = 0 '파일읽기 종료

                    OldVer = Split(str, "=")

                    If ReadOk2 = 1 And OldVer.Length > 1 Then '읽시 시작이 Ok면
                        If NewVer(0) = OldVer(0) And NewVer(1) = OldVer(1) Then
                            Download = False
                        End If
                    End If
                End While
                sr1.Close()

                If Download = True Then
                    If CheckFileName = "" Then
                        CheckFileName = NewVer(0) & "=" & NewVer(2)
                    Else
                        CheckFileName = CheckFileName & "/" & NewVer(0) & "=" & NewVer(2)
                    End If
                End If
            End If
        End While
        sr.Close()

        Label1.Text = "Version 비교완료."
        If CheckFileName = "" Then
            Label5.Text = "변경사항 없음."
            Application.DoEvents()
            Try
                'Shell(Application.StartupPath & "\YJ MMS.exe", AppWinStyle.NormalFocus)
                'CreateObject("WScript.Shell").Run(Chr(34) & Application.StartupPath & "\YJ Login.exe" & Chr(34), True)
                If UBound(Diagnostics.Process.GetProcessesByName("UpdateChecker")) < 0 Then
                    Shell(Application.StartupPath & "\UpdateChecker.exe", AppWinStyle.Hide)
                End If
                Shell(Application.StartupPath & "\YJ Login.exe", AppWinStyle.NormalFocus)
            Catch ex As Exception
                MsgBox("YJ Login.exe" & vbCrLf & ex.Message, MsgBoxStyle.Information, "프로그램 실행 오류")
            End Try
            Me.Close()
            System.IO.File.Delete(Application.StartupPath & "\NEWCheckVer.ini")
        Else
            FileDownLoad(ftpUrl & "/Run_File/")
            System.IO.File.Copy(Application.StartupPath & "\NEWCheckVer.ini", Application.StartupPath & "\CheckVer.ini", True)
            System.IO.File.Delete(Application.StartupPath & "\NEWCheckVer.ini")
        End If

    End Sub

    Private Sub FileDownLoad(ByVal ftpUrl As String)

        Dim j As Integer
        Dim File() As String
        File = Split(CheckFileName, "/")

        Label1.Text = "0 / " & File.Length & " EA Download중"

        Dim org_ftpUrl As String = ftpUrl

        For j = 0 To File.Length - 1

            Label2.Visible = True
            Label2.Text = File(j)
            Label1.Text = j + 1 & " / " & File.Length & " EA Download중"
            '############################  파일다운로드 관련부분 #########################################################
            ''''' 여기부분에 다운로드 경로를 설정할 수 있도록...................
            Dim localPath_download As String = String.Empty '다운로드 받을 경로

            Dim fileName As String = Split(File(j), "=")(0)
            Dim folder_name As String = Split(File(j), "=")(1)

            If folder_name = "ROOT" Then
                folder_name = String.Empty
                localPath_download = Application.StartupPath & "\"
                ftpUrl = org_ftpUrl
            Else
                localPath_download = Application.StartupPath & "\" & folder_name & "\"
                '폴더가 없으면 생성
                If System.IO.Directory.Exists(localPath_download) = False Then
                    System.IO.Directory.CreateDirectory(localPath_download)
                End If
                ftpUrl = org_ftpUrl & folder_name & "/"
            End If

            Console.WriteLine(localPath_download & fileName)
            Console.WriteLine(ftpUrl & fileName)

            Try
                Dim requestFileDownload As FtpWebRequest = Nothing
                Dim responseFileDownload As FtpWebResponse = Nothing

                requestFileDownload = DirectCast(WebRequest.Create(ftpUrl & fileName), FtpWebRequest)
                requestFileDownload.Credentials = New NetworkCredential(ftpID, ftpPassword)
                requestFileDownload.Method = WebRequestMethods.Ftp.DownloadFile
                requestFileDownload.UsePassive = True
                requestFileDownload.KeepAlive = False
                requestFileDownload.UseBinary = False
                responseFileDownload = DirectCast(requestFileDownload.GetResponse(), FtpWebResponse)

                Dim responseStream As Stream = responseFileDownload.GetResponseStream()

                Dim Length As Integer = 2048
                Dim buffer As [Byte]() = New [Byte](Length - 1) {}
                Dim bytesRead As Integer = responseStream.Read(buffer, 0, Length)

                '#####파일의 크기를 알아보기 위해...#####
                ''프로그래스바의 최대를 알아본다.
                Dim total_count As Integer = 0
                Dim run_Count As Integer = 0
                total_count = responseFileDownload.ContentLength
                'ProgressBar1.Maximum = responseFileDownload.ContentLength
                ProgressBar1.Maximum = 100
                ProgressBar1.Value = 0

                If fileName = "ERP Update.exe" Then '업데이트 파일이 Update.exe 파일이라면 파일면 변경
                    fileName = fileName & "_1"
                End If

                Dim writeStream As New FileStream(localPath_download & fileName, FileMode.Create)

                i = 2048

                While bytesRead > 0
                    writeStream.Write(buffer, 0, bytesRead)
                    bytesRead = responseStream.Read(buffer, 0, Length)
                    i += bytesRead
                    'ProgressBar1.Value += bytesRead '<--------왜?? 끝까지 안찰까? 맥시멈이랑 밸류랑 똑같은데
                    run_Count += bytesRead
                    ProgressBar1.Value = run_Count / total_count * 100
                    'Label2.Text = Split(File(j), "=")(0) & "    " & Format(ProgressBar1.Value, "#,##0") & " / " & Format(ProgressBar1.Maximum, "#,##0") & " byte"
                    Label2.Text = Split(File(j), "=")(0) & "    " & Format(run_Count, "#,##0") & " / " & Format(total_count, "#,##0") & " byte"

                    Application.DoEvents()
                End While
                writeStream.Close()

                requestFileDownload = Nothing
                responseFileDownload = Nothing

            Catch exceptionObj As Exception
                MsgBox(exceptionObj.Message.ToString(), MsgBoxStyle.Critical)
            End Try
            '############################  파일다운로드 관련부분 #########################################################
        Next j

        Label1.Text = "Update File Downloaded."
        Label5.Text = "Update 완료"
        'Shell(Application.StartupPath & "\YJ MMS.exe", AppWinStyle.NormalFocus)
        'CreateObject("WScript.Shell").Run(Chr(34) & Application.StartupPath & "\YJ Login.exe" & Chr(34), True)
        'MsgBox(UBound(Diagnostics.Process.GetProcessesByName("UpdateChecker")))
        If UBound(Diagnostics.Process.GetProcessesByName("UpdateChecker")) < 0 Then
            Shell(Application.StartupPath & "\UpdateChecker.exe", AppWinStyle.Hide)
        End If

        Shell(Application.StartupPath & "\YJ Login.exe", AppWinStyle.NormalFocus)

        Me.Close()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        If FTPDownloading = False Then
            FTPDownloading = True
            If Timer1.Enabled = True Then CheckVerDownload(ftpUrl & "/Run_File/", "NEWCheckVer.ini", True)
        End If

    End Sub

    '다른 다운로드 방법!!!!!!
    Private Sub FileDownLoad2(ByVal ftpUrl As String)

        Dim j As Integer
        Dim File() As String
        File = Split(CheckFileName, "/")

        Label1.Text = "0 / " & File.Length & " EA Download중"

        Dim org_ftpUrl As String = ftpUrl

        For j = 0 To File.Length - 1

            Label2.Visible = True
            Label2.Text = File(j)
            Label1.Text = j + 1 & " / " & File.Length & " EA Download중"
            '############################  파일다운로드 관련부분 #########################################################
            ''''' 여기부분에 다운로드 경로를 설정할 수 있도록...................
            Dim localPath_download As String = String.Empty '다운로드 받을 경로

            Dim fileName As String = Split(File(j), "=")(0)
            Dim folder_name As String = Split(File(j), "=")(1)

            If folder_name = "ROOT" Then
                folder_name = String.Empty
                localPath_download = Application.StartupPath & "\"
                ftpUrl = org_ftpUrl
            Else
                localPath_download = Application.StartupPath & "\" & folder_name & "\"
                '폴더가 없으면 생성
                If System.IO.Directory.Exists(localPath_download) = False Then
                    System.IO.Directory.CreateDirectory(localPath_download)
                End If
                ftpUrl = org_ftpUrl & folder_name & "/"
            End If

            Dim request As FtpWebRequest = CType(WebRequest.Create(ftpUrl & fileName), FtpWebRequest)
            request.Method = WebRequestMethods.Ftp.DownloadFile

            ' This example assumes the FTP site uses anonymous logon.
            request.Credentials = New NetworkCredential(ftpID, ftpPassword)

            Dim response As FtpWebResponse = CType(request.GetResponse(), FtpWebResponse)

            Dim responseStream As Stream = response.GetResponseStream()

            Dim Length As Integer = 1024
            Dim buffer As [Byte]() = New [Byte](Length - 1) {}
            Dim bytesRead As Integer = responseStream.Read(buffer, 0, Length)

            ProgressBar1.Maximum = response.ContentLength
            ProgressBar1.Value = 1

            If fileName = "ERP Update.exe" Then '업데이트 파일이 Update.exe 파일이라면 파일면 변경
                fileName = fileName & "_1"
            End If

            Dim writeStream As New FileStream(localPath_download & fileName, FileMode.Create)

            i = 1024

            While bytesRead > 0
                writeStream.Write(buffer, 0, bytesRead)
                bytesRead = responseStream.Read(buffer, 0, Length)
                i += bytesRead
                ProgressBar1.Value += bytesRead '<--------왜?? 끝까지 안찰까? 맥시멈이랑 밸류랑 똑같은데

                Label2.Text = Split(File(j), "=")(0) & "    " & Format(ProgressBar1.Value, "#,##0") & " / " & Format(ProgressBar1.Maximum, "#,##0") & " byte"

                Application.DoEvents()
            End While
            writeStream.Close()

            response.Close()

        Next j

        Label1.Text = "Update File Downloaded."
        Label5.Text = "Update 완료"
        'Shell(Application.StartupPath & "\YJ MMS.exe", AppWinStyle.NormalFocus)
        Shell(Application.StartupPath & "\YJ Login.exe", AppWinStyle.NormalFocus)

        Me.Close()

    End Sub
End Class

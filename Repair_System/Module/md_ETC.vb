Imports System.Globalization
Imports System.Threading

Module md_ETC

    '접속ID
    Public loginID As String = registryEdit.ReadRegKey("Software\Yujin\Repair System\Login", "User ID", "repair_system_user")

    Public Function GetWeekOfYear(ByVal targetDate As DateTime) As Integer
        Return GetWeekOfYear(targetDate, Nothing)
    End Function

    Public Function GetWeekOfYear(ByVal targetDate As DateTime, ByVal culture As CultureInfo) As Integer
        If culture Is Nothing Then culture = CultureInfo.CurrentCulture
        Dim weekRule As CalendarWeekRule = culture.DateTimeFormat.CalendarWeekRule
        Dim firstDayOfWeek As DayOfWeek = culture.DateTimeFormat.FirstDayOfWeek
        Return culture.Calendar.GetWeekOfYear(targetDate, weekRule, firstDayOfWeek)
    End Function

    Public Sub cellCal(sender As Object, e As MouseEventArgs)

        Dim tip As String = String.Empty
        Dim total_Qty As Double = 0
        Dim total_Number_Cell As Integer = 0
        Dim total_Char_Cell As Integer = 0
        Dim total_NumericCell As Integer = 0
        For i = sender.Selection.TopRow To sender.Selection.BottomRow
            For j = sender.Selection.LeftCol To sender.Selection.RightCol
                If IsNumeric(sender(i, j)) Then
                    total_Qty += sender(i, j)
                    total_NumericCell += 1
                End If
                If Not Trim(sender(i, j)) = String.Empty Then
                    If IsNumeric(sender(i, j)) Then
                        total_Number_Cell += 1
                    Else
                        total_Char_Cell += 1
                    End If
                End If
            Next
        Next

        If Not total_Char_Cell = 0 Then
            tip = "개수(문자) : " & Format(total_Char_Cell, "#,##0.########")
        End If

        If Not total_Number_Cell = 0 Then
            tip += "    개수(숫자) : " & Format(total_Number_Cell, "#,##0.########")
        End If

        If Not total_Qty = 0 Then
            tip += "    합계 : " & Format(total_Qty, "#,##0.########")
            tip += "    평균 : " & Format(total_Qty / total_NumericCell, "#,##0.########")
        End If

        frm_Main.lb_Status.Text = tip & "                                                              "

    End Sub

    Public th_LoadingWindow As Thread
    Dim thread_SleepTime As Integer = 300

    Public Sub thread_LoadingFormStart()

        th_LoadingWindow = New Thread(AddressOf load_LoadWindow)
        th_LoadingWindow.IsBackground = True
        th_LoadingWindow.SetApartmentState(ApartmentState.MTA)
        th_LoadingWindow.Start()
        Thread.Sleep(thread_SleepTime)

    End Sub

    Public Sub thread_LoadingFormStart(ByVal showText As String)

        th_LoadingWindow = New Thread(AddressOf load_LoadWindow2)
        th_LoadingWindow.IsBackground = True
        th_LoadingWindow.SetApartmentState(ApartmentState.MTA)
        th_LoadingWindow.Start(showText)
        Thread.Sleep(thread_SleepTime)

    End Sub

    Private Sub load_LoadWindow()

        frm_LoadingImage.ShowDialog()

    End Sub

    Private Sub load_LoadWindow2(ByVal showText As String)

        frm_LoadingImage.Label1.Text = showText
        frm_LoadingImage.ShowDialog()

    End Sub
End Module

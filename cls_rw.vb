Imports System.Net
Imports System.Data.OleDb
Public Class cls_rw
    Public Function getLastDateOfMonth(ByVal month As Integer, ByVal year As Integer) As Integer
        If month = 1 Or month = 3 Or month = 5 Or month = 7 Or month = 8 Or month = 10 Or month = 12 Then
            Return 31
        ElseIf month = 2 Then
            If year Mod 4 = 0 Then
                Return 29
            Else
                Return 28
            End If
        Else
            Return 30
        End If
    End Function

    Public Function get_ip_address() As String
        Dim c_ip_address As String = ""
        Dim myHost As String = Dns.GetHostName
        Dim ipEntry As IPHostEntry = Dns.GetHostEntry(myHost)

        For Each tmpIpAddress As IPAddress In ipEntry.AddressList
            If tmpIpAddress.AddressFamily = Sockets.AddressFamily.InterNetwork Then
                Dim ipAddress As String = tmpIpAddress.ToString
                c_ip_address = ipAddress
                Exit For
            End If
        Next
        Return c_ip_address
    End Function

    Public Function get_proc_date(ByVal k As OleDbConnection) As String
        Dim com As OleDbCommand
        Dim pd As String

        com = New OleDbCommand("select sysdate from dual", k)
        Dim tgl As DateTime = com.ExecuteScalar
        com.Dispose()
        pd = tgl.ToString("dd-MMM-yyyy")
        Return pd
    End Function

    Public Function get_proc_time(ByVal k As OleDbConnection) As String
        Dim com As OleDbCommand
        Dim pt As String

        com = New OleDbCommand("select sysdate from dual", k)
        Dim tgl As DateTime = com.ExecuteScalar
        com.Dispose()
        pt = tgl.ToString("HH:mm:ss")
        Return pt
    End Function

    Public Function w_date(ByVal k As OleDbConnection) As String
        Dim c As New cls_rw

        Dim wh As String = c.get_proc_date(kon)
        Dim js As String = c.get_proc_time(kon)
        Dim c_whdate As String = ""

        If js >= "16:00:00" Then
            c_whdate = DateAdd(DateInterval.Day, 1, CDate(c.get_proc_date(kon)))
            c_whdate = Format(CDate(c_whdate), "dd-MMM-yyyy")
        Else
            c_whdate = c.get_proc_date(kon)
        End If
        Return c_whdate

    End Function

    Public Function get_proc_date_plus(ByVal k As OleDbConnection) As String
        Dim com As OleDbCommand
        Dim pd As String

        com = New OleDbCommand("select sysdate+1 from dual", k)
        Dim tgl As DateTime = com.ExecuteScalar
        com.Dispose()
        pd = tgl.ToString("dd-MMM-yyyy")
        Return pd
    End Function

    Public Function wh_date(ByVal k As OleDbConnection) As String
        Dim c As New cls_rw

        Dim wh As String = c.get_proc_date(kon)
        Dim js As String = c.get_proc_time(kon)
        Dim c_whdate As String = ""

        If js >= "16:00:00" Then
            c_whdate = DateAdd(DateInterval.Day, 1, CDate(c.get_proc_date(kon)))
            c_whdate = Format(CDate(c_whdate), "dd-MMM-yyyy")
        Else
            c_whdate = c.get_proc_date(kon)
        End If
        Return c_whdate

    End Function

End Class

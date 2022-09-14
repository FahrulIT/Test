Imports System.Data.OleDb
Imports System.Configuration
Module mdl_rw
    Public kon As New OleDbConnection
    Public c_user_id As String
    Public _procNo As Integer = 0
    Public _recStatus As String = Nothing
    Public _whDetailNo As String = 0
    Public _cnDetailNo As String = 0
    Public _tahun As String = Nothing
    Public _bulan As String = Nothing
    Public _slipNo As String = Nothing
    Public _slipNoCancel As String = Nothing
    Public _genSlipNo As String = Nothing
    Public _genSlipNoCancel As String = Nothing
    Public _printName As String = Nothing

    Public Sub Koneksi()
        Try
            kon.Close()
            Dim ConString As String = ConfigurationManager.ConnectionStrings("ACTEM").ConnectionString
            kon = New OleDbConnection(ConString)
            kon.Open()
        Catch ex As Exception
            kon.Close()
        End Try
    End Sub

    Enum FormMode
        FormNew = 1
        FormEdit = 2
        FormDelete = 3
    End Enum

End Module
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Sql
Imports System.Data.Sqlclient
Imports System.Windows.Forms
Public Class Base

    Enum TipoDB
        MSSQL = 1
        MSAccess = 2
    End Enum
    Dim vNombre As String
    Dim vTipo As TipoDB
    Shared ConnAccess As OleDb.OleDbConnection
    Shared ConnSQL As Data.SqlClient.SqlConnection
    Public Function Conectar(ByVal TipoDeBase As TipoDB, ByVal RutaMDBoNombreBase As String) As Boolean
        'Conectar si fuera access
        Try
            vNombre = RutaMDBoNombreBase
            vTipo = TipoDB.MSAccess

            ConnAccess = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & RutaMDBoNombreBase)
            ConnAccess.Open()
            If ConnAccess.State = ConnectionState.Open Then
                System.Console.WriteLine("Conectado a DB" & " " & "MSAceess estado " & ConnAccess.State)
                Return True
            End If

        Catch ex As Exception
            MessageBox.Show("ERROR en Modulo de Datos: " & ex.Message & vbCrLf & " No se puede abrir la base de datos MSAcces, asegurese que exista el archivo")
            'System.Console.WriteLine(ex.Message)
        End Try
    End Function
    Public Function Conectar(ByVal TipoDeBase As TipoDB, ByVal RutaMDBoNombreBase As String, ByVal NombreServer As String, ByVal SeguridadIntegrada As Boolean) As Boolean
        'conectar sobrecargado si fuese MSSQLServer
        Try
            Dim SSI As String
            vNombre = RutaMDBoNombreBase
            vTipo = TipoDB.MSSQL
            If SeguridadIntegrada = True Then
                SSI = "yes"
            Else
                SSI = "no"
            End If

            ConnSQL = New Data.SqlClient.SqlConnection("Server=" & NombreServer & "; " & "database=" & RutaMDBoNombreBase & "; integrated security=" & SSI & ";")
            ConnSQL.Open()
            If ConnSQL.State = ConnectionState.Open Then
                System.Console.WriteLine("Conectado a DB" & " " & "MSSQLserver estado " & ConnSQL.State)
                Return True
            End If

        Catch ex As Exception
            System.Console.WriteLine(ex.Message)
        End Try
    End Function
    Public Function TomarDatos(ByVal SQLString As String) As DataTable

        Dim dt As DataTable
        dt = New DataTable
        Try
            If vTipo = TipoDB.MSAccess Then
                Dim da As OleDb.OleDbDataAdapter
                da = New OleDb.OleDbDataAdapter(SQLString, ConnAccess)
                da.Fill(dt)
            End If

            If vTipo = TipoDB.MSSQL Then
                Dim da As SqlClient.SqlDataAdapter
                da = New SqlClient.SqlDataAdapter(SQLString, ConnSQL)
                da.Fill(dt)
            End If

        Catch ex As Exception
            System.Console.WriteLine(ex.Message)
        End Try
        Return dt
    End Function
    Public Function TomarSchema() As DataTable
        Dim dt As DataTable
        dt = New DataTable
        If vTipo = TipoDB.MSAccess Then
            dt = ConnAccess.GetSchema()
        End If

        If vTipo = TipoDB.MSSQL Then
            dt = ConnSQL.GetSchema
        End If

        Return dt
    End Function
    Public Function TomarFechaHora() As String
        Dim dt As DataTable
        dt = New DataTable
        Try
            If vTipo = TipoDB.MSAccess Then
                Dim da As OleDb.OleDbDataAdapter
                da = New OleDb.OleDbDataAdapter("select getdate()", ConnAccess)
                da.Fill(dt)
            End If

            If vTipo = TipoDB.MSSQL Then
                Dim da As SqlClient.SqlDataAdapter
                da = New SqlClient.SqlDataAdapter("select getdate()", ConnSQL)
                da.Fill(dt)
            End If

        Catch ex As Exception
            System.Console.WriteLine(ex.Message)
        End Try
        Return CStr(dt.Rows(0).Item(0))

    End Function
    Public Function ExecutarQuery(ByVal SQLstring As String) As Int32
        If vTipo = TipoDB.MSAccess Then
            ' Public Function ExecuteStatement(ByVal strSQL As String) As Int32
            Dim intRecCount As Int32 = -1
            Try
                '  Dim objConn As New OleDbConnection(strConnection)
                ' objConn.Open()
                Dim objCommand As New OleDb.OleDbCommand
                With objCommand
                    .CommandType = CommandType.Text
                    .CommandText = SQLstring
                    .Connection = ConnAccess
                    intRecCount = .ExecuteNonQuery()
                End With
                '     objConn.Close()
                Return intRecCount
            Catch ex As Exception
                Return -1
            End Try
            'End Function
        End If

        If vTipo = TipoDB.MSSQL Then
            'Public Function ExecuteStatement(ByVal strSQL As String) As Int32
            Dim intRecCount As Int32 = -1
            Try

                '    objConn.Open()
                Dim objCommand As New SqlCommand
                With objCommand
                    .CommandType = CommandType.Text
                    .CommandText = SQLstring
                    .Connection = ConnSQL
                    intRecCount = .ExecuteNonQuery()
                End With
                '  objConn.Close()
                Return intRecCount
            Catch ex As Exception
                Return -1
            End Try
            '    End Function
        End If

    End Function
    Public ReadOnly Property Nombre()
        Get
            Return vNombre
        End Get
    End Property
    Public ReadOnly Property TipoBaseDatos()
        Get
            Return vTipo
        End Get
    End Property
    Public Sub New()

    End Sub
End Class

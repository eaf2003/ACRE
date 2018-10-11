
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Sql
Imports System.Data.Sqlclient
Imports System.Windows.Forms
Public Class BaseAccess
    Implements iBase
    Enum TipoDB
        MSSQL = 1
        MSAccess = 2
    End Enum
    Dim vNombre As String
    Dim vTipo As TipoDB
    Shared Conn As OleDb.OleDbConnection
    Function Conectar(ByVal RutaMDBoNombreBase As String) As Boolean Implements iBase.Conectar

        'Conectar si fuera access
        Try
            vNombre = RutaMDBoNombreBase
            vTipo = TipoDB.MSAccess

            Conn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & RutaMDBoNombreBase)
            Conn.Open()
            If Conn.State = ConnectionState.Open Then
                System.Console.WriteLine("Conectado a DB" & " " & "MSAceess estado " & Conn.State)
                Return True
            End If

        Catch ex As Exception
            MessageBox.Show("ERROR en Modulo de Datos: " & ex.Message & vbCrLf & " No se puede abrir la base de datos MSAcces, asegurese que exista el archivo")
            'System.Console.WriteLine(ex.Message)
        End Try
    End Function
    Function Conectar(ByVal RutaMDBoNombreBase As String, ByVal NombreServer As String, ByVal SeguridadIntegrada As Boolean) As Boolean Implements iBase.Conectar
        Return False
    End Function

    Function TomarDatos(ByVal SQLString As String) As DataTable Implements iBase.TomarDatos

        Dim dt As DataTable
        dt = New DataTable
        Try

            Dim da As OleDb.OleDbDataAdapter
            da = New OleDb.OleDbDataAdapter(SQLString, Conn)
            da.Fill(dt)




        Catch ex As Exception
            System.Console.WriteLine(ex.Message)
        End Try
        Return dt
    End Function
    Function TomarSchema() As DataTable Implements iBase.TomarSchema
        Dim dt As DataTable
        dt = New DataTable

        dt = Conn.GetSchema()



        Return dt
    End Function
    Function TomarFechaHora() As String Implements iBase.TomarFechaHora
        Dim dt As DataTable
        dt = New DataTable
        Try

            Dim da As OleDb.OleDbDataAdapter
            da = New OleDb.OleDbDataAdapter("select getdate()", Conn)
            da.Fill(dt)


        Catch ex As Exception
            System.Console.WriteLine(ex.Message)
        End Try
        Return CStr(dt.Rows(0).Item(0))

    End Function
    Function ExecutarQuery(ByVal SQLstring As String) As Int32 Implements iBase.ExecutarQuery

        ' Public Function ExecuteStatement(ByVal strSQL As String) As Int32
        'devuelvo cantidad de regs tratados
        Dim intRecCount As Int32 = -1
        Try
            '  Dim objConn As New OleDbConnection(strConnection)
            ' objConn.Open()
            Dim objCommand As New OleDb.OleDbCommand
            With objCommand
                .CommandType = CommandType.Text
                .CommandText = SQLstring
                .Connection = Conn
                'intRecCount = .ExecuteNonQuery()
                intRecCount = .ExecuteScalar()
            End With
            '     objConn.Close()
            Return intRecCount
        Catch ex As Exception
            Return -1
        End Try
        'End Function


    End Function
    Public Function CantRegs(ByVal SQLstring As String) As Int32 Implements iBase.CantRegs
        Dim dt As New DataTable
        Dim o As New OleDb.OleDbDataAdapter(SQLstring, Conn)
        o.Fill(dt)
        Return dt.Rows.Count
    End Function

    Public ReadOnly Property Nombre() Implements iBase.Nombre
        Get
            Return vNombre
        End Get
    End Property
    Public ReadOnly Property TipoBaseDatos() Implements iBase.TipoBaseDatos
        Get
            Return vTipo
        End Get
    End Property
    Public Sub New()
        '   DB1 = New Datos.BaseAccess
        Dim RutaBase = System.AppDomain.CurrentDomain.BaseDirectory & "itorg.mdb"
        Conectar(RutaBase)
    End Sub
End Class

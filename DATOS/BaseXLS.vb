Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Sql
Imports System.Data.Sqlclient
Imports System.Windows.Forms
Public Class BaseXLS
    Implements iBase
    Enum TipoDB
        MSSQL = 1
        MSAccess = 2
    End Enum
    Dim vNombre As String
    Dim vTipo As TipoDB
    Shared Conn As OleDb.OleDbConnection
    Private da As OleDb.OleDbDataAdapter
    Private dt As DataTable = New DataTable
    Private BindingSource1 As New BindingSource()
    Private SQLSelCommand1 As New OleDbCommand
    Private Builder1 As New OleDbCommandBuilder
    Function Conectar(ByVal RutaXLS As String) As Boolean Implements iBase.Conectar

        Try
            'Uso un builder porque no se deja pasar el filename string...
            Dim vXLSfile As String = RutaXLS
            Dim connStringBuilder As New OleDbConnectionStringBuilder()

            connStringBuilder.DataSource = CStr(vXLSfile)
            connStringBuilder.Provider = "Microsoft.ACE.OLEDB.12.0"
            connStringBuilder.Add("Extended Properties", "Excel 8.0;HDR=Yes;IMEX=1")

            Conn = New OleDbConnection(connStringBuilder.ConnectionString)
            Conn.Open()


        Catch ex As Exception
            MessageBox.Show("ERROR en Modulo de Datos: " & ex.Message & vbCrLf & _
                            "connstring:" & Conn.ConnectionString & vbCrLf & _
            "RutaXLS:" & RutaXLS & vbCrLf & _
                            " No se puede abrir la base de datos MSAcces, asegurese que exista el archivo")
        End Try
    End Function
    Function Conectar(ByVal RutaMDBoNombreBase As String, ByVal NombreServer As String, ByVal SeguridadIntegrada As Boolean) As Boolean Implements iBase.Conectar
        Return False
    End Function

    Function TomarDatos(ByVal SQLString As String) As DataTable Implements iBase.TomarDatos

        '        dt = New DataTable
        Try


            da = New OleDb.OleDbDataAdapter(SQLString, Conn)
            da.Fill(dt)


        Catch ex As Exception
            System.Console.WriteLine(ex.Message)
        End Try
        Return dt
    End Function
    Function TomarDatos(ByVal SQLString As String, ByRef ControlBind As Object) As DataTable 'Implements iBase.TomarDatos

        '        dt = New DataTable
        Try
            'creo el sqlcommand
            SQLSelCommand1.Connection = Conn
            SQLSelCommand1.CommandText = SQLString

            'instancio el DAdapter
            '            da = New OleDb.OleDbDataAdapter(SQLString, Conn)
            da = New OleDb.OleDbDataAdapter
            da.SelectCommand = SQLSelCommand1

            'lleno el commandbuilder
            Builder1.DataAdapter = da


            da.Fill(dt)

            '           dt.PrimaryKey("Nº Rva") = dt.Columns("Nº Rva")

            '            dt.Columns.Add(New DataColumn("Nº Rva"))
            Dim PrimaryKey(0) As DataColumn
            Dim KeyCol As New DataColumn()
            KeyCol.ColumnName = "Nº Rva"
            PrimaryKey(0) = KeyCol
            PrimaryKey(0) = dt.Columns(1)
            dt.PrimaryKey = PrimaryKey
            '          primaryKey(1) = dt.Columns("Nº Rva")
            '         dt.PrimaryKey = primaryKey
            System.Console.WriteLine("PK:" & dt.PrimaryKey(0).ColumnName)
            ControlBind.DataSource = dt

            '            BindingSource1.DataSource = dt
            '           ControlBind.DataSource = BindingSource1

        Catch ex As Exception
            System.Console.WriteLine(ex.Message)
        End Try
        Return dt
    End Function

    Sub UpdateDatosBindeados()
        '        Dim da As New OleDbDataAdapter

        '        updcommand.CommandText = updcommand

        '       da.UpdateCommand = 1123


        ' Dim myBuilder As OleDbCommandBuilder = New OleDbCommandBuilder(da)

        ' myBuilder.GetUpdateCommand()

        'da.UpdateCommand = myBuilder.GetUpdateCommand()
        System.Console.WriteLine("PK:" & dt.PrimaryKey(0).ColumnName)

        da.Update(dt)

        '        System.Console.WriteLine(CStr(ControlBindeado.DataSource))
    End Sub

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
        'Dim RutaBase = System.AppDomain.CurrentDomain.BaseDirectory & "itorg.mdb"
        'Conectar(RutaBase)
    End Sub
End Class

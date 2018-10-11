Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Sql
Imports System.Data.Sqlclient
Imports System.Windows.Forms
Imports System.Threading
Imports System.Globalization
Public Class BaseSQL
    Implements iBase
    Enum TipoDB
        MSSQL = 1
        MSAccess = 2
    End Enum
    Dim vNombreBase As String
    Dim vOrigenDatos As String
    Dim vTipo As TipoDB
    '   Dim vConfig As New conf
    Shared ConnAccess As OleDb.OleDbConnection
    Public Shared ConnSQL As Data.SqlClient.SqlConnection
    Dim Configuracion As New Configuracion
    Dim da As New SqlClient.SqlDataAdapter
    Dim dt As New Data.DataTable
    '  Dim Builder1 As New SqlCommandBuilder
    Dim SQLSelCommand1 As New SqlCommand
    Public Sub ShowCurrentCulture()
        MsgBox("Culture of {0} in application domain {1}: {2}" & Thread.CurrentThread.Name & AppDomain.CurrentDomain.FriendlyName & CultureInfo.CurrentCulture.Name & CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern)

    End Sub

    Public Function Conectar(ByVal RutaMDBoNombreBase As String) As Boolean Implements iBase.Conectar
        Return False
    End Function
    Public Function Conectar(ByVal RutaMDBoNombreBase As String, ByVal NombreServer As String, ByVal SeguridadIntegrada As Boolean) As Boolean Implements iBase.Conectar
        'conectar sobrecargado si fuese MSSQLServer
        '  ShowCurrentCulture()
        Try
            Dim SSI As String
            vNombreBase = RutaMDBoNombreBase
            vTipo = TipoDB.MSSQL
            If SeguridadIntegrada = True Then
                SSI = "true"
                ConnSQL = New Data.SqlClient.SqlConnection("data source=" & NombreServer & "; " & "initial catalog=" & RutaMDBoNombreBase & "; integrated security=" & SSI & ";")

            Else
                SSI = "false"
                ConnSQL = New Data.SqlClient.SqlConnection("data source=" & NombreServer & "; " & "initial catalog=" & RutaMDBoNombreBase & ";user id=" & Configuracion.DevolverValor("user", "base.xml") & ";password =" & Configuracion.DevolverValor("password", "base.xml"))
            End If

            ConnSQL.Open()
            vOrigenDatos = ConnSQL.DataSource
            If ConnSQL.State = ConnectionState.Open Then
                System.Console.WriteLine("Conectado a DB" & " " & "MSSQLserver estado " & ConnSQL.State)
                Return True
            End If

        Catch ex As Exception
            MsgBox("Error en BaseSQL.Conectar: " & ConnSQL.ConnectionString & " ->MSG: " & ex.Message)

        End Try
    End Function
    Public Function TomarDatos(ByVal SQLString As String) As DataTable Implements iBase.TomarDatos

        'Dim dt As DataTable
        'dt = New DataTable
        Try
            If vTipo = TipoDB.MSAccess Then
                Dim da As OleDb.OleDbDataAdapter
                da = New OleDb.OleDbDataAdapter(SQLString, ConnAccess)
                da.Fill(dt)
            End If

            If vTipo = TipoDB.MSSQL Then
                '                Dim da As SqlClient.SqlDataAdapter

                '               da = New SqlClient.SqlDataAdapter(SQLString, ConnSQL)

                SQLSelCommand1.Connection = ConnSQL
                SQLSelCommand1.CommandText = SQLString
                'instancio el DAdapter
                da.SelectCommand = SQLSelCommand1

                System.Console.WriteLine(SQLString)
                'dt.Rows.Clear()
                da.Fill(dt)



            End If

        Catch ex As Exception
            MsgBox(SQLString & " - " & ex.Message)
        End Try
        Return dt
    End Function

    Public Function TomarDatosNewDT(ByVal SQLString As String) As DataTable

        Dim dt As DataTable
        dt = New DataTable
        Try
            If vTipo = TipoDB.MSAccess Then
                Dim da As OleDb.OleDbDataAdapter
                da = New OleDb.OleDbDataAdapter(SQLString, ConnAccess)
                da.Fill(dt)
            End If

            If vTipo = TipoDB.MSSQL Then
                '                Dim da As SqlClient.SqlDataAdapter

                '               da = New SqlClient.SqlDataAdapter(SQLString, ConnSQL)

                SQLSelCommand1.Connection = ConnSQL
                SQLSelCommand1.CommandText = SQLString
                'instancio el DAdapter
                da.SelectCommand = SQLSelCommand1

                System.Console.WriteLine(SQLString)
                'dt.Rows.Clear()
                da.Fill(dt)



            End If

        Catch ex As Exception
            MsgBox(SQLString & " - " & ex.Message)
        End Try
        Return dt
    End Function
    Function TomarDatos(ByVal SQLString As String, ByVal PrimaryKName As String) As DataTable 'Implements iBase.TomarDatos
        'ESTE ESTA PARA PRUEBAS O DISCONTINUAR A FUTURO EAF
        '        dt = New DataTable
        Try
            'creo el sqlcommand
            SQLSelCommand1.Connection = ConnSQL
            SQLSelCommand1.CommandText = SQLString
            'instancio el DAdapter
                  da.SelectCommand = SQLSelCommand1
    

            da.Fill(dt)

            '           dt.PrimaryKey("Nº Rva") = dt.Columns("Nº Rva")

            '            dt.Columns.Add(New DataColumn("Nº Rva"))
            ''accccaaa COMO definir una PK en un dt o ds
            'Dim PrimaryKey(0) As DataColumn
            'Dim KeyCol As New DataColumn()
            'KeyCol.ColumnName = PrimaryKName
            'PrimaryKey(0) = KeyCol
            'PrimaryKey(0) = dt.Columns(PrimaryKName)
            'dt.PrimaryKey = PrimaryKey
            ''accacaaaa()

            '          primaryKey(1) = dt.Columns("Nº Rva"
            '         dt.PrimaryKey = primaryKey
            ' System.Console.WriteLine("PK:" & dt.PrimaryKey(0).ColumnName)
            '  ControlBind.DataSource = dt

            '            BindingSource1.DataSource = dt
            '           ControlBind.DataSource = BindingSource1

        Catch ex As Exception
            System.Console.WriteLine(ex.Message)
        End Try
        Return dt
    End Function
    Public Sub RefrescarBase() 'sincroniza todo de nuevo
        dt.Clear()
        da.Fill(dt)

    End Sub
    Public Sub cleardt()
        dt.Clear()

    End Sub

    Public Sub UpdateDatosBindeados() 'sirve para modificaciones sincroniza dt con db
        'para que le builder cree los comandos de update e insert la tabla de dt debe tener una PK y ser SQL
        'lleno el commandbuilder
        Try

            Dim builder1 As New SqlCommandBuilder(da)
            '        Console.WriteLine(builder1.GetUpdateCommand().CommandText)
            da.Update(dt)

        Catch ex As Exception
            MsgBox("BaseSQL.UpdataDatosBindeados " & ex.Message)
        End Try
    End Sub


    Public Function TomarSchema() As DataTable Implements iBase.TomarSchema
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
    Public Function TomarFechaHora() As String Implements iBase.TomarFechaHora
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
    Public Function ExecutarQuery(ByVal SQLstring As String) As Int32 Implements iBase.ExecutarQuery
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

                Return intRecCount
            Catch ex As Exception

                MsgBox("BasSQL.ExecutarQuery:  " & SQLstring & "MSG-> " & ex.Message)
                Return -1
            End Try

        End If

    End Function
    Public Sub ExecutarQuery(ByRef SqlCmdObj As SqlCommand)
        SqlCmdObj.Connection = ConnSQL
        SqlCmdObj.ExecuteNonQuery()

    End Sub
    Public Function CantRegs(ByVal SQLstring As String) As Int32 Implements iBase.CantRegs
        Dim dt As New DataTable
        Dim o As New SqlDataAdapter(SQLstring, ConnSQL)
        o.Fill(dt)

        'dt.Dispose()
        'o.Dispose()

        Return dt.Rows.Count
    End Function
    Public ReadOnly Property NombreOrigen()
        Get
            Return vOrigenDatos
        End Get
    End Property
    Public ReadOnly Property Nombre() Implements iBase.Nombre
        Get
            Return vNombreBase
        End Get
    End Property
    Public ReadOnly Property TipoBaseDatos() Implements iBase.TipoBaseDatos
        Get
            Return vTipo
        End Get
    End Property

    '  Public ReadOnly Property Conexion() Implements iBase.Conexion
    '     Get
    '         Return Me.ConnSQL
    '     End Get
    '  End Property

    Public Sub New()
        
        'DB1 = New Datos.BaseSQL
        InitCultura()
        '  Console.WriteLine(Configuracion.DevolverValor("base"))
        '   Conectar(vConfig.base, vConfig.server, vConfig.ssi)
        ' Conectar(Configuracion.DevolverValor("base", Configuracion.DevolverValor("server"), Configuracion.DevolverValor("ssi"))

    End Sub
    Public Sub New(ByVal dbname As String, ByVal server As String, ByVal ssi As Boolean)
        'DB1 = New Datos.BaseSQL

        '  Console.WriteLine(Configuracion.DevolverValor("base"))
        '    Conectar(vConfig.base, vConfig.server, vConfig.ssi)
        Conectar(dbname, server, ssi)

    End Sub

    Protected Overrides Sub Finalize()
        'ConnSQL.Close()
        MyBase.Finalize()
    End Sub

End Class
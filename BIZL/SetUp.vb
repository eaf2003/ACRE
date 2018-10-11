Imports System.Threading
Imports System.Globalization
Public Class SetUp
    Private ObaseSetupSQL As New Datos.BaseSQL
    Private oCfg As New Datos.Configuracion
    Private oDTSetup As New DataTable
    Public oDTMoneda As New DataTable
    Dim oDTFactura As New DataTable
    Dim oDTTCredito As New DataTable
    Dim oDTTReserva As New DataTable

    Dim vFechaCena As Date
    Public Sub LoadSetup()

    End Sub

    Public ReadOnly Property Moneda() As DataTable
        Get
            Return oDTMoneda
        End Get
    End Property

    Public ReadOnly Property TipoReserva() As DataTable
        Get
            Return oDTTReserva
        End Get
    End Property

    Public ReadOnly Property TCredito() As DataTable
        Get
            Return oDTTCredito
        End Get
    End Property

    Public ReadOnly Property TipoFactura() As DataTable
        Get
            Return oDTFactura
        End Get
    End Property

    Public Property FechaCena() As Date
        Get
            Return vFechaCena
        End Get
        Set(ByVal value As Date)

        End Set
    End Property


    Public Sub New()
        Try

            InitCultura()

            ObaseSetupSQL.Conectar(oCfg.DevolverValor("base", "base.xml"), oCfg.DevolverValor("server", "base.xml"), oCfg.DevolverValor("ssi", "base.xml"))
            Dim qRy As String = "select * FROM Reservassetup"
            oDTSetup = ObaseSetupSQL.TomarDatos(qRy)
            vFechaCena = oDTSetup.Rows(0)("FECHA")

            '            ShowCurrentCulture()

            'ObaseSetupSQL.cleardt()
            Dim qry2 As String = "SELECT * FROM TIPOCAMBIO"
            oDTMoneda = ObaseSetupSQL.TomarDatos(qry2)

            'ObaseSetupSQL.cleardt()

            Dim qRy3 As String = "SELECT * FROM TIPOFACTURA"
            oDTFactura = ObaseSetupSQL.TomarDatos(qRy3)

            '           ObaseSetupSQL.cleardt()

            Dim qRy4 As String = "SELECT * FROM TIPOTARJETAsCREDITO"
            oDTTCredito = ObaseSetupSQL.TomarDatos(qRy4)

            '          ObaseSetupSQL.cleardt()

            Dim qRy5 As String = "SELECT * FROM TIPORESERVA"
            oDTTReserva = ObaseSetupSQL.TomarDatos(qRy5)


        Catch ex As Exception
            MsgBox("Class Setup.New: ->>" & ex.Message)

        End Try

    End Sub
End Class

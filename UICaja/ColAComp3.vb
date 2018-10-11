Public Class ColAComp3
    Private vreserva As Integer
    Private vnombre As String
    Private vingreso As Boolean
    Private vacredito As Integer
    Private vcantpax As Integer
    Private vPago As Double
    Private vPrecioUnit As Double
    Private vabona As Boolean
    'CLASE PARA SOPORTAR LA COLECCION DE MAS DE 2 VARIABLES, LA USO PARA BUSCAR POR NOMBRE INGRESO ID, Y CARGAR LA LISTA CON ESTOS VALORES
    Public Property IDReserva() As Integer
        Get
            Return vreserva
        End Get
        Set(ByVal value As Integer)
            vreserva = value
        End Set
    End Property
    Public Property Nombre() As String
        Get
            Return vnombre
        End Get
        Set(ByVal value As String)
            vnombre = value
        End Set
    End Property
    Public Property abona() As Boolean
        Get
            Return vabona
        End Get
        Set(ByVal value As Boolean)
            vabona = value
        End Set
    End Property

    Public Property Ingreso() As Boolean
        Get
            Return vingreso
        End Get
        Set(ByVal value As Boolean)
            vingreso = value
        End Set
    End Property

    Public Property Acredito() As Integer
        Get
            Return vacredito
        End Get
        Set(ByVal value As Integer)
            vacredito = value
        End Set
    End Property

    Public Property CantPax() As Integer
        Get
            Return vcantpax
        End Get
        Set(ByVal value As Integer)
            vcantpax = value
        End Set
    End Property

    Public Property pago() As Double
        Get
            Return vPago
        End Get
        Set(ByVal value As Double)
            vPago = value
        End Set
    End Property

    Public Property PrecioUnit() As Double
        Get
            Return vPrecioUnit
        End Get
        Set(ByVal value As Double)
            vPrecioUnit = value
        End Set
    End Property

End Class

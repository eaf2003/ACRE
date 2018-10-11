Imports Microsoft.SqlServer
Imports ACRE.Datos
Imports System.Data

Public Class Reserva
    Inherits Reservas
    Public nombre As String
    'NO LLAMAR DIRECTO A ESA CLASSE, HACERLO ATRAVEZ DE RESERVAS CON S, QUE ES UNA COLLECCION DE RESERVA,S
    '  Private oCfg As New Datos.Configuracion
    '   Private oBase As New Datos.BaseSQL
    '    Private odtReserva As New DataTable
    Private vIDreserva As Integer
    Private vMesa As Integer
    Private vCantPax As Integer
    Private vNombre As String
    Private vServicio As String
    Private vTipoReserva As String
    Private vPrecioUnit As Double
    Private vTipoFactura As String
    Private vMozoID As Integer
    Private vIngreso As Boolean
    Private vImprimioT As Integer
    Private vObservaciones As String
    Private vTV As Integer
    Private vTI As Integer
    Private vClaseTango As Boolean
    Private vPago As Double

    Private vPrePago As Double
    Private RowFound() As DataRow

    Public Sub Acreditar()
        '        Dim filaAcre As DataRow
        '       filaAcre = odtReserva.Select(IDReserva = vIDreserva)
        '   RowFound(0)("PrecioUnit")
        RowFound(0)("ImprimioToken") = vImprimioT + 1
        RowFound(0)("FechaImp") = oBase.TomarFechaHora
        '  RowFound(0)("Ingreso") = True
        Me.GuardarCambios()
    End Sub
    Public Sub Ingresar()
        If Not (IsDBNull(RowFound(0)("ImprimioToken"))) Then
            If RowFound(0)("ImprimioToken") > 0 Then
                RowFound(0)("Ingreso") = True
                RowFound(0)("FechaIngreso") = oBase.TomarFechaHora
                Me.GuardarCambios()
                Me.SincronizarBase()
            End If
        Else
            MsgBox("Reserva.Ingresar: No puede ingresar una reserva que no se acredito, primero acreditar luego Ingresar")
        End If
    End Sub
    Public Property ImprimioToken() As Integer
        Get
            Return vImprimioT
        End Get
        Set(ByVal value As Integer)

        End Set
    End Property
    Public Property Ingreso() As Boolean
        Get
            Return vIngreso
        End Get
        Set(ByVal value As Boolean)

        End Set
    End Property
    Public Property ClaseTango() As Boolean
        Get
            Return vClaseTango
        End Get
        Set(ByVal value As Boolean)
            RowFound(0)("ClaseTango") = value
            Me.GuardarCambios()
        End Set
    End Property
    Public Property MozoID() As Integer
        Get
            Return vMozoID
        End Get
        Set(ByVal value As Integer)

        End Set
    End Property

    Public Property TipoFactura() As String
        Get
            Return vTipoFactura
        End Get
        Set(ByVal value As String)
            RowFound(0)("TipoFactura") = value
            Me.GuardarCambios()
        End Set
    End Property

    Public Property PrecioUnitario() As Double
        Get
            Return vPrecioUnit
        End Get
        Set(ByVal value As Double)
            RowFound(0)("preciounit") = value
            Me.GuardarCambios()
        End Set
    End Property

    Public Property Pago() As Double
        Get
            Return vPago
        End Get
        Set(ByVal value As Double)
            RowFound(0)("pago") = value
            Me.GuardarCambios()
        End Set
    End Property

    Public Property TipoReserva() As String
        Get
            Return vTipoReserva
        End Get
        Set(ByVal value As String)
            RowFound(0)("TipoReserva") = value
            Me.GuardarCambios()
        End Set
    End Property

    Public Property Servicio() As String
        Get
            Return vServicio
        End Get
        Set(ByVal value As String)
            RowFound(0)("Servicio") = value
            Me.GuardarCambios()
        End Set
    End Property


    Public Property NombreReserva() As String
        Get
            Return vNombre
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Property Observaciones() As String
        Get
            Return vObservaciones
        End Get
        Set(ByVal value As String)
            RowFound(0)("observaciones") = value
            Me.GuardarCambios()
        End Set
    End Property
    Private Sub GetReserva()

    End Sub

    Public ReadOnly Property IDReserva() As Integer
        Get
            Return vIDreserva
        End Get

    End Property

    Public Property Mesa() As Integer
        Get
            Return vMesa
        End Get
        Set(ByVal value As Integer)
            RowFound(0)("mesa") = value
            '     vMesa = value
            Me.GuardarCambios()
        End Set
    End Property
    Public Property CantidadPax() As Integer
        Get
            Return vCantPax
        End Get
        Set(ByVal value As Integer)
            RowFound(0)("cantpax") = value
            Me.GuardarCambios()
        End Set
    End Property

    Public Property TransferVuelta() As Integer
        Get
            Return vTV
        End Get
        Set(ByVal value As Integer)
            RowFound(0)("TransferV") = value
            Me.GuardarCambios()
        End Set
    End Property

    Public Property TransferIda() As Integer
        Get
            Return vTI
        End Get
        Set(ByVal value As Integer)
            RowFound(0)("Transferi") = value
            Me.GuardarCambios()
        End Set
    End Property

    Public Property Prepago() As Integer
        Get
            Return vPrePago
        End Get
        Set(ByVal value As Integer)
            RowFound(0)("Prepago") = value
            Me.GuardarCambios()
        End Set
    End Property

    Public Sub Crear(ByVal IDReserva As Integer, ByVal CantidadPax As Short, ByVal Mesa As Short, ByVal PrecioUnit As Integer, ByVal MozoID As Integer, ByVal Notas As String, ByVal NombreReserva As String, ByVal Servicio As String, ByVal TipoReserva As String, ByVal ModoPago As String, ByVal Pagado As Integer, ByVal TransferID As Integer)

    End Sub

    Public Sub Pagar(ByVal IDReserva As Integer, ByVal Cantidad As Integer, ByVal Moneda As String)

    End Sub

    Public Sub Eliminar(ByVal IDReserva As Integer)

    End Sub

    Public Sub AsignarMozo(ByVal IDReserva As Integer, ByVal MozoID As Integer)

    End Sub

    Public Sub TotalDeuda(ByVal IDReserva As Integer)

    End Sub

    Public Sub Upgradear(ByVal Upgrade As String, ByVal IDReserva As Integer)

    End Sub

    Public Sub DarPropina(ByVal IDReserva As Integer, ByVal IDMozo As Integer)

    End Sub

    Public Sub New()

    End Sub

    Public Sub New(ByVal FechaCena As Date, ByVal idreserva As Integer)
        ' oBase.Conectar(oCfg.DevolverValor("base", "acreditador.xml"), oCfg.DevolverValor("server", "acreditador.xml"), oCfg.DevolverValor("ssi", "acreditador.xml"))
        'odtReserva = oBase.TomarDatos("select * from reservas where idreserva=" & idreserva & "and fecha=" & FechaCena)
    End Sub

    Public Sub New(ByVal aidreserva As Integer)
        '    oBase.Conectar(oCfg.DevolverValor("base", "acreditador.xml"), oCfg.DevolverValor("server", "acreditador.xml"), oCfg.DevolverValor("ssi", "acreditador.xml"))
        '   odtReserva = oBase.TomarDatos("select * from reservas where idreserva=" & idreserva)

        'BUSCO EN EL DT QUE YA TENGO CARGADO LA ROW QUE TIENE EL ID DE RESERVA Y POPULATE LA CLASE RESRVA
        Try
            Dim consulta As String
            consulta = "IDReserva =" & aidreserva

            '          Dim rowfound() As DataRow
            rowfound = odtReserva.Select(consulta)
            If (rowfound.Count = 0) Then
                MsgBox("class reserva: no encontro la reserva segun el id, aca deria llegar siempre un ID, algo anda mal EAF")
                Return
            End If

            vIDreserva = aidreserva
            vNombre = If(IsDBNull(rowfound(0)("Nombre")), 0, rowfound(0)("Nombre"))
            vMesa = If(IsDBNull(rowfound(0)("mesa")), 0, rowfound(0)("mesa")) 'devuelve un valor solo el qry asi que es un 
            vCantPax = If(IsDBNull(rowfound(0)("CantPax")), 0, rowfound(0)("CantPax"))
            vPrecioUnit = RowFound(0)("PrecioUnit")
            vPrePago = RowFound(0)("PrePago")
            vPago = If(IsDBNull(RowFound(0)("pago")), 0, RowFound(0)("pago"))
            vServicio = RowFound(0)("Servicio")
            vTipoFactura = rowfound(0)("TipoFactura")
            vTipoReserva = rowfound(0)("TipoReserva")
            vObservaciones = rowfound(0)("Observaciones")
            vMozoID = If(IsDBNull(rowfound(0)("MozoID")), 0, rowfound(0)("MozoID"))
            vIngreso = If(IsDBNull(rowfound(0)("Ingreso")), False, rowfound(0)("Ingreso"))
            vImprimioT = If(IsDBNull(rowfound(0)("ImprimioToken")), 0, rowfound(0)("ImprimioToken"))
            vTV = If(IsDBNull(RowFound(0)("TransferV")), 0, RowFound(0)("TransferV"))
            vTI = If(IsDBNull(RowFound(0)("TransferI")), 0, RowFound(0)("TransferI"))
            vClaseTango = If(IsDBNull(RowFound(0)("clasetango")), False, RowFound(0)("clasetango"))

        Catch ex As Exception
            MsgBox("Class Reserva: " & IDReserva & "->>" & ex.Message)

        End Try
        'odtReserva.Rows(0)("mesa") 'devuelve un valor solo el qry asi que es un 

    End Sub



End Class

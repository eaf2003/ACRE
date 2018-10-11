Public Class Cobros
    Private ObaseSQL As New Datos.BaseSQL
    Private oCfg As New Datos.Configuracion
    Private oDTCobros As DataTable
    Private oDTCobroReserva As New DataTable
    Private RowFound() As DataRow
    Private vFechacena As Date


    Public Sub Cobrar(ByVal aIDReserva As Integer, ByVal aMonto As Double, ByVal ATipoCambio As Double, ByVal aMedio As String, ByVal aIDMoneda As String, ByVal aDESC As Double)
        Dim RowNew As DataRow
        Dim vDtimeServer As DateTime
        Try

            vDtimeServer = ObaseSQL.TomarFechaHora
            RowNew = oDTCobros.NewRow

            RowNew("ReservaID") = aIDReserva
            RowNew("Monto") = aMonto
            RowNew("MonedaID") = aIDMoneda
            RowNew("MonedaTC") = ATipoCambio
            RowNew("fecha") = vDtimeServer
            RowNew("Medio") = aMedio
            RowNew("descuento") = aDESC
            RowNew("fechacena") = vfechacena

            oDTCobros.Rows.Add(RowNew)
            ObaseSQL.UpdateDatosBindeados()


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub New(ByVal afechaCenaL As Date)
        ObaseSQL.Conectar(oCfg.DevolverValor("base", "base.xml"), oCfg.DevolverValor("server", "base.xml"), oCfg.DevolverValor("ssi", "base.xml"))
        Dim qRy As String = "select * FROM cobros where (DATEDIFF(day, Fechacena,'" & afechaCenaL & "') = 0)"
        oDTCobros = ObaseSQL.TomarDatos(qRy)
        vfechacena = afechaCenaL
    End Sub
    Public Function VerCobroReserva(ByVal aidreserva As Integer) As DataTable
        oDTCobroReserva.Rows.Clear()
        '      busco si pago la reserva en el DT actual , si no pago no hago el qry a la base
        '  Dim consulta As String
        'consulta = "ReservaID =" & aidreserva
        'RowFound = oDTCobros.Select(consulta)
        'If Not (RowFound.Count = 0) Then

        '    Dim qRy As String = "select * FROM cobros where ReservaID=" & aidreserva
        '    oDTCobroReserva = ObaseSQL.TomarDatos(qRy)
        'End If


        Dim qRy As String = "select * FROM cobros where ReservaID=" & aidreserva
        ObaseSQL.cleardt()
        oDTCobroReserva = ObaseSQL.TomarDatos(qRy)

        Return oDTCobroReserva

    End Function

    Public Sub UpdateadDatosBindeados()
        ObaseSQL.UpdateDatosBindeados()
    End Sub
    Public ReadOnly Property TotalARS() As Double
        Get
            Dim vtotal As Double
            Dim consulta As String
            consulta = "sum(Monto)"
            Dim filtro As String
            filtro = "MonedaID='ARS' AND medio='EFECTIVO'"

            Dim vObjtCompute As Object
            vObjtCompute = oDTCobros.Compute(consulta, filtro)

            vtotal = If(IsDBNull(vObjtCompute), 0, CDbl(vObjtCompute.ToString))

            Return vtotal

        End Get
    End Property

    Public ReadOnly Property TotalUSD() As Double
        Get
            Dim vtotal As Double
            Dim consulta As String
            consulta = "sum(Monto)"
            Dim filtro As String
            filtro = "MonedaID='USD' AND medio='EFECTIVO'"

            Dim vObjtCompute As Object
            vObjtCompute = oDTCobros.Compute(consulta, filtro)

            vtotal = If(IsDBNull(vObjtCompute), 0, CDbl(vObjtCompute.ToString))

            Return vtotal

        End Get
    End Property
    Public ReadOnly Property TotalBRL() As Double
        Get
            Dim vtotal As Double
            Dim consulta As String
            consulta = "sum(Monto)"
            Dim filtro As String
            filtro = "MonedaID='BRL' AND medio='EFECTIVO'"

            Dim vObjtCompute As Object
            vObjtCompute = oDTCobros.Compute(consulta, filtro)

            vtotal = If(IsDBNull(vObjtCompute), 0, CDbl(vObjtCompute.ToString))

            Return vtotal

        End Get
    End Property
    Public ReadOnly Property TotalEUR() As Double
        Get
            Dim vtotal As Double
            Dim consulta As String
            consulta = "sum(Monto)"
            Dim filtro As String
            filtro = "MonedaID='EUR' AND medio='EFECTIVO'"

            Dim vObjtCompute As Object
            vObjtCompute = oDTCobros.Compute(consulta, filtro)

            vtotal = If(IsDBNull(vObjtCompute), 0, CDbl(vObjtCompute.ToString))

            Return vtotal

        End Get
    End Property

    Public ReadOnly Property TotalTARJETA() As Double
        Get
            Dim vtotal As Double
            Dim consulta As String
            consulta = "sum(Monto)"
            Dim filtro As String
            filtro = "MonedaID='ARS' AND medio <> 'EFECTIVO'"

            Dim vObjtCompute As Object
            vObjtCompute = oDTCobros.Compute(consulta, filtro)

            vtotal = If(IsDBNull(vObjtCompute), 0, CDbl(vObjtCompute.ToString))

            Return vtotal

        End Get
    End Property



End Class

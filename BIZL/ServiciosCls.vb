Public Class ServiciosCls
    Private ObaseSQL As New Datos.BaseSQL
    Private oCfg As New Datos.Configuracion
    Private oDTServicios As New DataTable
    Public Sub LoadServicios(ByVal aorderBy)
        ObaseSQL.Conectar(oCfg.DevolverValor("base", "base.xml"), oCfg.DevolverValor("server", "base.xml"), oCfg.DevolverValor("ssi", "base.xml"))
        Dim qRy As String = "select * FROM Tiposervicio ORDER BY " & aorderBy
        oDTServicios = ObaseSQL.TomarDatos(qRy)

        '       odtReserva = (oBase.TomarDatos("select * FROM Reservas order by IDreserva desc"))
    End Sub

    Public Sub New()
        LoadServicios("servicio")
    End Sub

    ReadOnly Property TiposServicio() As DataTable
        Get
            Return oDTServicios
        End Get
    End Property
End Class

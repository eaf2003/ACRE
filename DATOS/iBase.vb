Public Interface iBase
    Function Conectar(ByVal RutaMDBoNombreBase As String) As Boolean

    Function Conectar(ByVal RutaMDBoNombreBase As String, ByVal NombreServer As String, ByVal SeguridadIntegrada As Boolean) As Boolean

    Function TomarDatos(ByVal SQLString As String) As DataTable

    Function TomarSchema() As DataTable
 
    Function TomarFechaHora() As String

    Function ExecutarQuery(ByVal SQLstring As String) As Int32

    Function CantRegs(ByVal SQLstring As String) As Int32

    ReadOnly Property Nombre()

    '    ReadOnly Property Conexion()

    ReadOnly Property TipoBaseDatos()
End Interface

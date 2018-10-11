Imports Microsoft.SqlServer
Imports ACRE.Datos
Imports System.Data


Public Class Reservas
    'heredo las funciones de las colleciones de net
    Inherits System.Collections.CollectionBase
    'instancio 1 vez sola la base y el file de config
    Public Shared oCfg As New Datos.Configuracion
    Public Shared oBase As New Datos.BaseSQL
    Public Shared odtReserva As New DataTable
    Private vFechaCena As Date
    Private vOrderByField As String
    Public Function AgregarWalkIn(ByVal aNombre As String, ByVal aServicio As String, ByVal atipoReserva As String, ByVal aMesa As Integer, ByVal aCantidad As Integer, ByVal aPrecioUnit As Double, ByVal aTipoFact As String, ByVal aObservaciones As String) As Integer
        Dim RowNew As DataRow
        Dim IDcalcWI As Integer

        Try

            LoadReservas(vFechaCena)
            'recargo de nuevo el dt de reservas asi me aseguro que me da el ultimo ID
            RowNew = odtReserva.NewRow 'creo la row que luego agregare al dt
            ' IDcalcWI = 7000000 + CInt(odtReserva.Rows(odtReserva.Rows.Count / 2).Item("IDReserva")) + odtReserva.Rows.Count + 1 'este ID sera autogenerado, no se por las dudas pongo el idactual + 7000000 + la cantidad de rows + 1 , asi no tengo chance de repetir el IDrESERVA y que guarde, agarro un nro del medio de la lista y le sumo la cuenta esa asi tengo menos chance de duplicar

            'saco el id nuevo ,busco de no duplicar la reserva y que sea un nro por arriba de 7m
            'defino que todos los WIs seran por arriba del ID 7millon
            Dim IDmasAlto As Integer = 0
            For Each fila In odtReserva.Rows 'recorro las reservas abiertas
                If fila("IDReserva") > IDmasAlto Then ' busco id mas alto
                    IDmasAlto = fila("IDReserva")
                    IDcalcWI = IDmasAlto + 1 ' seteo el id con con el id mas alto
                    If IDmasAlto < 7000000 Then 'si no hay ningun WI antes crea uno nuevo
                        IDcalcWI = IDmasAlto + 7000000 + 1
                    End If
                End If

            Next

            'Dim consulta As String
            'Dim rowdup() As DataRow
            'consulta = "IDReserva =" & IDcalcWI
            'rowdup = odtReserva.Select(consulta)
            'If (rowdup.Count > 0) Then
            '    MsgBox("mismo ID")
            '    IDcalcWI = IDcalcWI + 1

            'End If
            'fin sacar ID

            RowNew("IDReserva") = IDcalcWI '7000000 + CInt(odtReserva.Rows(0).Item("IDReserva")) + odtReserva.Rows.Count + 1 'este ID sera autogenerado, no se por las dudas pongo el idactual + 7000000 + la cantidad de rows + 1 , asi no tengo chance de repetir el IDrESERVA y que guarde
            RowNew("Fecha") = vFechaCena
            RowNew("Nombre") = aNombre
            RowNew("Servicio") = aServicio
            RowNew("TipoReserva") = atipoReserva
            RowNew("Mesa") = aMesa
            RowNew("CantPax") = aCantidad
            RowNew("PrecioUnit") = aPrecioUnit
            RowNew("TipoFactura") = aTipoFact
            RowNew("Observaciones") = aObservaciones
            RowNew("ImprimioToken") = 1
            RowNew("FechaImp") = oBase.TomarFechaHora
            RowNew("preciounit") = 1
            RowNew("prepago") = 0

            '        RowNew("Nombre") =
            odtReserva.Rows.Add(RowNew)
            'odtReserva.AcceptChanges()
            Me.GuardarCambios() 'guardo los cambios en la base
            '  odtReserva.Dispose()
            LoadReservas(vFechaCena) ' recargo el dt de nuevo, no se si sirve
        Catch ex As Exception
            MsgBox("error al agregar reserva " & IDcalcWI & " - " & ex.Message)
        End Try
        Return IDcalcWI
    End Function
    ReadOnly Property DTReserva() As DataTable
        Get
            Return odtReserva
        End Get
    End Property

    'Public Function Listar(ByVal QryDT As String) As DataRow

    'End Function

    Public Sub GuardarCambios()
        oBase.UpdateDatosBindeados()
    End Sub

    Public Sub Add(ByVal areserva As Reserva)
        ' List.Add(areserva)

    End Sub
    Public Sub Remove(ByVal index As Integer)
        ' Check to see if there is a widget at the supplied index.
        If index > Count - 1 Or index < 0 Then
            ' If no widget exists, a messagebox is shown and the operation is 
            ' cancelled.
            MsgBox("Index not valid!")
        Else
            ' Invokes the RemoveAt method of the List object.
            List.RemoveAt(index)
        End If
    End Sub

    Public ReadOnly Property Item(ByVal idreserva As Integer) As Reserva
        Get
            ' The appropriate item is retrieved from the List object and 
            ' explicitly cast to the Widget type, then returned to the 
            ' caller.
            Dim res As New Reserva(idreserva)
            Return res

            '    Return CType(List.Item(index), Reserva)
        End Get
    End Property

    Public Sub New()
        ' oBase.Conectar(oCfg.DevolverValor("base", "acreditador.xml"), oCfg.DevolverValor("server", "acreditador.xml"), oCfg.DevolverValor("ssi", "acreditador.xml"))
        ' odtReserva = oBase.TomarDatos("select * from reservas")
    End Sub

    Public Sub New(ByVal aFechaCena As Date) 'esto deberia ser para que devulva listado
        LoadReservas(aFechaCena)
    End Sub
    Public Sub New(ByVal aFechaCena As Date, ByVal aOrderBYfield As String) 'esto deberia ser para que devulva listado
        LoadReservas(aFechaCena, aOrderBYfield)
    End Sub
    Public Sub New(ByVal aFechaCena As Date, ByVal aWhere As String, ByVal aOrderBYfield As String) 'esto deberia ser para que devulva listado
        LoadReservas(aFechaCena, aWhere, aOrderBYfield)
    End Sub
    Public Sub LoadReservas(ByVal aFechaCenaL As DateTime)
        Try
            odtReserva.Clear()
            vFechaCena = aFechaCenaL
            oBase.Conectar(oCfg.DevolverValor("base", "acreditador.xml"), oCfg.DevolverValor("server", "acreditador.xml"), oCfg.DevolverValor("ssi", "acreditador.xml"))
            '        Dim qRy As String = "select * FROM Reservas WHERE (DATEDIFF(day, Fecha,'" & aFechaCenaL & "') = 0) ORDER BY nombre"
            Dim qRy As String = "select * FROM Reservas WHERE (DATEDIFF(day, Fecha,'" & FormatDateTime(aFechaCenaL, DateFormat.ShortDate) & "') = 0) ORDER BY nombre"

            odtReserva = oBase.TomarDatos(qRy)
            '       odtReserva = (oBase.TomarDatos("select * FROM Reservas order by IDreserva desc"))
        Catch ex As Exception
            MsgBox("Error en Reservas.LoadReservas - " & ex.Message)
        End Try

    End Sub
    Public Sub LoadReservas(ByVal aFechaCenaL As Date, ByVal aOrderBYfield As String)
        Try
            odtReserva.Clear()
            vFechaCena = aFechaCenaL
            vOrderByField = aOrderBYfield
            oBase.Conectar(oCfg.DevolverValor("base", "acreditador.xml"), oCfg.DevolverValor("server", "acreditador.xml"), oCfg.DevolverValor("ssi", "acreditador.xml"))
            Dim qRy As String = "select * FROM Reservas WHERE (DATEDIFF(day, Fecha,'" & aFechaCenaL & "') = 0) ORDER BY " & aOrderBYfield
            odtReserva = oBase.TomarDatos(qRy)


        Catch ex As Exception
            MsgBox("Error en Reservas.LoadReservas - " & ex.Message)
        End Try
        '       odtReserva = (oBase.TomarDatos("select * FROM Reservas order by IDreserva desc"))
    End Sub
    Public Sub LoadReservas(ByVal aFechaCenaL As Date, ByVal aWhereClause As String, ByVal aOrderBYfield As String)
        Try

            odtReserva.Clear()
            vFechaCena = aFechaCenaL
            vOrderByField = aOrderBYfield
            oBase.Conectar(oCfg.DevolverValor("base", "acreditador.xml"), oCfg.DevolverValor("server", "acreditador.xml"), oCfg.DevolverValor("ssi", "acreditador.xml"))
            Dim qRy As String = "select * FROM Reservas WHERE " & aWhereClause & " AND (DATEDIFF(day, Fecha,'" & aFechaCenaL & "') = 0) ORDER BY " & aOrderBYfield
            odtReserva = oBase.TomarDatos(qRy)

            '       odtReserva = (oBase.TomarDatos("select * FROM Reservas order by IDreserva desc"))

        Catch ex As Exception
            MsgBox("Error en Reservas.LoadReservas - " & ex.Message)
        End Try


    End Sub
    Public Sub SincronizarBase()
        oBase.RefrescarBase()

    End Sub
End Class

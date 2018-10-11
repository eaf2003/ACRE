Imports System.Threading
Imports System.Globalization
Public Class CenaPasada
    Dim oXLSFILE As New Datos.BaseXLS
    Dim oBaseSql As New Datos.BaseSQL
    Dim dtsql As New DataTable
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        '--------------------------------------------------------------
        oBaseSql.Conectar(Form1.cfg.DevolverValor("base", "base.xml"), Form1.cfg.DevolverValor("server", "base.xml"), Form1.cfg.DevolverValor("ssi", "base.xml"))
        '        Dim qry As String = "select * from reservas where fecha='" & Form1.DateTimePicker2.Value.Date & "' order by idreserva desc"
        Dim qry As String = "SELECT Reservas.IDreserva, Reservas.Fecha, Reservas.Nombre, Reservas.Servicio, Reservas.TipoReserva, Reservas.Mesa, Reservas.TipoFactura, Reservas.CantPax, Reservas.PrecioUnit," _
& "Reservas.PrePago, CAST(Reservas.Observaciones AS varchar(MAX)) AS Observaciones, Reservas.IDtabla, Reservas.MozoID, Reservas.Ingreso, Reservas.ImprimioToken, Reservas.FechaImp," _
& "Reservas.TransferI, Reservas.TransferV, Reservas.FechaIngreso, Reservas.ClaseTango, Reservas.pago, SUM(Cobros_1.Monto) AS ARS," _
& " (SELECT SUM(Monto) AS Expr1" _
& " FROM Cobros " _
 & " WHERE (ReservaID = Reservas.IDreserva) AND (MonedaID = 'USD') AND (Medio = 'EFECTIVO')" _
  & " GROUP BY MonedaID) AS USD," _
   & " (SELECT SUM(Monto) AS Expr1" _
    & " FROM Cobros AS Cobros_4" _
     & " WHERE (ReservaID = Reservas.IDreserva) AND (MonedaID = 'EUR') AND (Medio = 'EFECTIVO')" _
      & " GROUP BY MonedaID) AS EUR," _
       & " (SELECT SUM(Monto) AS Expr1" _
        & " FROM Cobros AS Cobros_3" _
         & " WHERE (ReservaID = Reservas.IDreserva) AND (MonedaID = 'BRL') AND (Medio = 'EFECTIVO')" _
          & " GROUP BY MonedaID) AS BRL," _
           & " (SELECT SUM(Monto) AS EXPR1" _
            & " FROM Cobros AS Cobros_2" _
             & " WHERE (ReservaID = Reservas.IDreserva) AND (MonedaID = 'ARS') AND (Medio <> 'EFECTIVO')" _
              & " GROUP BY MonedaID) AS TARJETA" _
& " FROM Reservas LEFT OUTER JOIN" _
 & " Cobros AS Cobros_1 ON Reservas.IDreserva = Cobros_1.ReservaID AND Cobros_1.MonedaID = 'ARS' AND Cobros_1.Medio = 'EFECTIVO'" _
 & " WHERE reservas.fecha='" & Form1.DateTimePicker2.Value.Date & "'" _
& " GROUP BY Reservas.IDreserva, Reservas.Fecha, Reservas.Nombre, Reservas.Servicio, Reservas.TipoReserva, Reservas.Mesa, Reservas.TipoFactura, Reservas.CantPax, Reservas.PrecioUnit, " _
  & "Reservas.PrePago, Reservas.IDtabla, Reservas.MozoID, Reservas.Ingreso, Reservas.ImprimioToken, Reservas.FechaImp, Reservas.TransferI, Reservas.TransferV, Reservas.FechaIngreso, " _
   & "Reservas.ClaseTango, Reservas.pago, CAST(Reservas.Observaciones AS varchar(MAX))"

        '        MsgBox(Form1.DateTimePicker2.Value.Date)

        dtsql = oBaseSql.TomarDatos(qry)
        DataGridView1.DataSource = dtsql

        ' Add any initialization after the InitializeComponent() call.
        Try
            'SET FECHA

            'If MsgBox("DEFINIR FECHA DE LA CENA DE HOY COMO:" & Form1.DateTimePicker1.Text, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            '    Dim sqlCmd2 As New SqlClient.SqlCommand
            '    sqlCmd2.Parameters.Add(New SqlClient.SqlParameter("@Fecha", SqlDbType.SmallDateTime))
            '    Dim strSQLUPDCmd2 As String = "UPDATE RESERVASSETUP SET FECHA=@Fecha WHERE PARAMETRO='FECHACENA' "
            '    sqlCmd2.CommandText = strSQLUPDCmd2

            '    sqlCmd2.Parameters("@Fecha").Value = Form1.DateTimePicker1.Value.Date
            '    oBaseSql.ExecutarQuery(sqlCmd2)

            'Else
            '    Return

            'End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Migrador_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Splitter1_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs)

    End Sub

    Private Sub OpenToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub HelpToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Extras.Click





    End Sub

    Private Sub Salir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Salir.Click
        Form1.Dispose()
        Me.Dispose()

    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guardar.Click
        'MsgBox("hola", MsgBoxStyle.Exclamation)
        oBaseSql.UpdateDatosBindeados()
        DataGridView1.EndEdit() 'sin esto no updatea el ultimo valor editado
    End Sub
End Class

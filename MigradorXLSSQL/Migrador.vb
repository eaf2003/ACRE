Imports System.Threading
Imports System.Globalization
Public Class Migrador
    Dim oXLSFILE As New Datos.BaseXLS
    Dim oBaseSql As New Datos.BaseSQL
    Dim dtsql As New DataTable
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

         '--------------------------------------------------------------
        oXLSFILE.Conectar(Form1.TextBox1.Text)
        'abro la basesql , y parametros de id ,tabla a migrar
        oBaseSql.Conectar(Form1.cfg.DevolverValor("base", "migrador.xml"), Form1.cfg.DevolverValor("server", "migrador.xml"), Form1.cfg.DevolverValor("ssi", "migrador.xml"))
        Dim tablaMigra As String = Form1.cfg.DevolverValor("TablaMigra", "migrador.xml")
        Dim tablaCampoId As String = Form1.cfg.DevolverValor("TablaCampoID", "migrador.xml")
        '        Dim dtSQL as new datatable = obase

        'abro el xls siempre que id reserva exista y sus parametros
        Dim strHoja As String = Form1.TextBox2.Text
        Dim strCampoID As String = Form1.TextBox3.Text
        '   Dim sqlXLS As String = "select * from [" & strHoja & "] where [" & strCampoID & "] is not null"
        Dim sqlXLS As String = "select * from [" & strHoja & "] where [" & strCampoID & "] is not null"
        Dim dtXLS As New DataTable
        dtXLS = oXLSFILE.TomarDatos(sqlXLS)

        'nombre de columnas del Excel
        Dim strXLSPrePago, strXLSTI, strXLSTV, strXLSIDreserva, strXLSFecha, strXLSNombre, strXLSServicio, strXLSTipoReserva, strXLSMesa, strXLSCantPax, strXLSPrecioUnit, strXLSTipoFactura, strXLSObservaciones As String

        Dim strSQLinsertCmd As String
        Dim sqlCmd1 As New SqlClient.SqlCommand

        'genero el insert parametrizo porque si no hay problemas con tipo de datos y comimillas y etc
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@IDreserva", SqlDbType.Int, 10))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@Fecha", SqlDbType.SmallDateTime))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@Nombre", SqlDbType.NVarChar, 100))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@Servicio", SqlDbType.NVarChar, 100))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@TipoReserva", SqlDbType.NVarChar, 100))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@Mesa", SqlDbType.SmallInt))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@CantPax", SqlDbType.SmallInt))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@PrecioUnit", SqlDbType.Float))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@TipoFactura", SqlDbType.NVarChar, 100))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@Observaciones", SqlDbType.Text))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@TI", SqlDbType.Int, 10))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@TV", SqlDbType.Int, 10))
        sqlCmd1.Parameters.Add(New SqlClient.SqlParameter("@PrePago", SqlDbType.Float))

        strSQLinsertCmd = "INSERT INTO " & tablaMigra _
& "(IDreserva,Fecha,Nombre,Servicio,TipoReserva,Mesa,CantPax,PrecioUnit,TipoFactura,Observaciones,transferi,transferv,prepago) VALUES (@IDreserva,@Fecha,@Nombre,@Servicio,@TipoReserva,@Mesa,@CantPax,@PrecioUnit,@TipoFactura,@Observaciones,@TI,@TV,@PrePago)"

        sqlCmd1.CommandText = strSQLinsertCmd


        '    System.Console.WriteLine(dtXLS.Rows(0).Item(dtXLS.Columns(0)))
        '        System.Console.WriteLine(dtXLS.Rows(0).Item(0).ToString())

        Try

            Dim CantRegsdelXLS As Integer = 0
            Dim CantAgregados As Integer = 0
            'recorro y busco en basesql si no esta esa reserva la agrego, ver a futuro de usar solo DT offline seria mas rapido y seguro en vez de escribir directo
            For Each rFila As DataRow In dtXLS.Rows 'recorre cada row, r tiene la fila
                CantRegsdelXLS = CantRegsdelXLS + 1
                For iCol As Integer = 0 To rFila.ItemArray.Count - 1 'celda x celda ,recorre cada columbra de la row r 'i va a tener el nro de columna
                    System.Console.WriteLine(rFila(iCol)) 'tira el valor de la celda de iColX de las rFilaX Y dtXLS.Columns(0).ColumnName da el nombre de la col segun el indice
                Next

                If (oBaseSql.CantRegs("select * from " & tablaMigra & " where " & tablaCampoId & "=" & rFila.Item(strCampoID) & "")) = 0 Then                 'si no existe el nro de reserva en la base inserto la reserva
                    CantAgregados = CantAgregados + 1

                    'strXLSIDreserva = rFila(strCampoID)
                    strXLSFecha = Form1.DateTimePicker1.Value.Date ' ponerle la fecha del server db asi es puro
                    'strXLSNombre = rFila("Nombre de Reserva").ToString()
                    strXLSServicio = rFila("Serv").ToString()
                    'strXLSTipoReserva = rFila("Tipo").ToString()
                    strXLSMesa = rFila("MESA").ToString()
                    strXLSCantPax = rFila("CUB").ToString()
                    strXLSPrecioUnit = rFila("$UN").ToString()
                    strXLSTI = rFila("TI").ToString
                    strXLSTV = rFila("TV").ToString
                    'strXLSTipoFactura = rFila(16).ToString() 'el xls tiene la col sin nombre por eso pongo el index, "col-14?"
                    'strXLSObservaciones = rFila("Observaciones").ToString()
                    strXLSPrePago = rFila("Pre").ToString()

                    sqlCmd1.Parameters("@IDreserva").Value = rFila(strCampoID)
                    sqlCmd1.Parameters("@Fecha").Value = strXLSFecha
                    sqlCmd1.Parameters("@Nombre").Value = rFila("Nombre de Reserva").ToString()
                    sqlCmd1.Parameters("@Servicio").Value = If(strXLSServicio.Trim.Length > 0, strXLSServicio, "EAFAIL")
                    sqlCmd1.Parameters("@TipoReserva").Value = rFila("Tipo").ToString()
                    sqlCmd1.Parameters("@Mesa").Value = If(strXLSMesa.Trim.Length > 0, strXLSMesa, "000")
                    sqlCmd1.Parameters("@CantPax").Value = If(strXLSCantPax.Trim.Length > 0, strXLSCantPax, "000")
                    sqlCmd1.Parameters("@PrecioUnit").Value = If(strXLSPrecioUnit.Trim.Length > 0, strXLSPrecioUnit, "0")
                    sqlCmd1.Parameters("@TipoFactura").Value = rFila("TipoFactura").ToString()
                    sqlCmd1.Parameters("@Observaciones").Value = rFila("Observaciones").ToString()
                    sqlCmd1.Parameters("@TI").Value = If(strXLSTI.Trim.Length > 0, strXLSTI, "0")
                    sqlCmd1.Parameters("@TV").Value = If(strXLSTV.Trim.Length > 0, strXLSTV, "0")
                    sqlCmd1.Parameters("@PrePago").Value = If(strXLSPrePago.Trim.Length > 0, strXLSPrePago, "0")



                    '            strSQLinsert = "INSERT INTO " & tablaMigra _
                    '       & "(IDreserva,Fecha,Nombre,Servicio,TipoReserva,Mesa,CantPax,PrecioUnit,TipoFactura,Observaciones) VALUES (" & _
                    'strXLSIDreserva & "," & strXLSFecha & "," & strXLSNombre & "," & strXLSServicio & "," & strXLSTipoReserva & "," & strXLSMesa & "," & strXLSCantPax & "," & strXLSPrecioUnit & "," & strXLSTipoFactura & "," & strXLSObservaciones & ")"
                    '          
                    'ejecutar el insert
                    oBaseSql.ExecutarQuery(sqlCmd1)
                    System.Console.WriteLine(rFila(strCampoID))
                End If

            Next
            'oBaseSql.ExecutarQuery(

            If CantAgregados = 0 Then
                MsgBox("no se agrego nada. de un XLS con " & CantRegsdelXLS & vbCrLf & "Si debio agregearse, Verifique:" + _
     vbCrLf + "1-Excel con encabezado, 2-archivo XLS correcto, 3-los registros con mismo nro de reserva, no existan previamente", MsgBoxStyle.OkOnly)
            Else
                MsgBox("Se procesaron:" & CantAgregados & " de " & CantRegsdelXLS & " Registros de XLS", MsgBoxStyle.OkOnly)
            End If


            ' DataGridView1.DataSource = oXLSFILE.TomarDatos(sqlXLS, DataGridView1)

            '   Dim dtsql As New DataTable


            dtsql = oBaseSql.TomarDatos("select * from reservas where fecha='" & Form1.DateTimePicker1.Value.Date & "'")
            DataGridView1.DataSource = dtsql


            ' Add any initialization after the InitializeComponent() call.


            'SET FECHA

            If MsgBox("DEFINIR FECHA DE LA CENA DE HOY COMO:" & Form1.DateTimePicker1.Text, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                'Dim SQLQRYFECHA As String = "UPDATE RESERVASSETUP SET FECHA='" & Form1.DateTimePicker1.Value.Date & "' WHERE PARAMETRO='FECHACENA' "
                'oBaseSql.ExecutarQuery(SQLQRYFECHA)
                Dim sqlCmd2 As New SqlClient.SqlCommand
                sqlCmd2.Parameters.Add(New SqlClient.SqlParameter("@Fecha", SqlDbType.SmallDateTime))
                Dim strSQLUPDCmd2 As String = "UPDATE RESERVASSETUP SET FECHA=@Fecha WHERE PARAMETRO='FECHACENA' "
                sqlCmd2.CommandText = strSQLUPDCmd2

                sqlCmd2.Parameters("@Fecha").Value = Form1.DateTimePicker1.Value.Date
                '  sqlCmd2.ExecuteNonQuery()
                oBaseSql.ExecutarQuery(sqlCmd2)

            Else
                Return

            End If

        Catch ex As Exception
            MsgBox("Error: " & ex.Message & vbCrLf & sqlCmd1.CommandText & vbCrLf)
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

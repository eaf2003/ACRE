Imports System.Threading
Imports System.Globalization
Public Class Caja
    Public Shared oReservas As BIZL.Reservas
    Public Shared oSetup As BIZL.SetUp
    Public IDLista
    Public Shared FechaReservas As Date

    Public CantResAcreditadas As Integer
    Public CantIngresados As Integer
    Public TotalReservas As Integer
    Public CantPersAcreditadas As Integer
    Public CantPersIngresadas As Integer
    Public TotalPers As Integer
    Public CantResPagadas As Integer


    Dim cambioFicha As Integer = 0
    Dim queCambio() As String
    Dim queCambioCol As Dictionary(Of String, Integer)

    Public oCobros As BIZL.Cobros

    'de ficha
    Dim vCantPax
    Dim vPrecioUnit
    Dim vTipoFactura
    Dim vPrepago
    Dim vServicio
    Dim vMesa
    Dim vMozo
    Dim vTipoCli
    Dim vTI
    Dim vTv
    Dim vObs
    Dim vCTango
    Dim vPago
    Dim vIngreso
    Dim vAcredito

    Private Sub CambiarLabelColorFG(ByVal controls As ControlCollection, ByVal fgColor As Color)
        If controls Is Nothing Then Return
        For Each C As Control In controls
            If TypeOf C Is Label Then DirectCast(C, Label).ForeColor = fgColor
            If C.HasChildren Then CambiarLabelColorFG(C.Controls, fgColor)
        Next
    End Sub
    Public Sub CargarListaSinBuscar(ByVal ListViewRes As ListView)
        ListViewRes.Items.Clear() 'fijarse de darle clear a los items y no a toda la lista, sino borro los headers tambien
        Dim str1(2) As String
        For Each s As KeyValuePair(Of Int32, ColAComp3) In AcomP3 ' busco s en la coleccion acomp2 , s tiene dos componentes value y key, key bindea al id y valua nombre
            'para llenar el listview
            str1(0) = s.Key.ToString
            str1(1) = s.Value.Nombre
            str1(2) = s.Value.abona

            Dim LVi1 As New ListViewItem(str1)
            'le meto el color si ya ingreso la persona en cuestion


            If ((s.Value.Ingreso = True And s.Value.Acredito > 0)) Then 'ingreso y acredito y no pago
                LVi1.BackColor = Color.LightSkyBlue
            End If

            If ((s.Value.Ingreso <> True And s.Value.Acredito > 0)) Then 'acredito pero no ingreso
                LVi1.BackColor = Color.WhiteSmoke
            End If

            If ((s.Value.pago > 0) Or s.Value.abona = True) Then 'pago, no importa la condicion
                LVi1.BackColor = Color.LightGreen
            End If




            ListViewRes.Items.Add(LVi1)
        Next
        ListViewRes.Columns(1).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)


        'act indicadores
        ToolStripProgressBar1.Maximum = ListView1.Items.Count
        ActulizarIndicadores()

    End Sub

    Public Sub CargarListaSinBuscarAcred(ByVal ListViewRes As ListView)
        ListViewRes.Items.Clear() 'fijarse de darle clear a los items y no a toda la lista, sino borro los headers tambien
        Dim str1(2) As String
        For Each s As KeyValuePair(Of Int32, ColAComp3) In AcomP3 ' busco s en la coleccion acomp2 , s tiene dos componentes value y key, key bindea al id y valua nombre
            'para llenar el listview
            str1(0) = s.Key.ToString
            str1(1) = s.Value.Nombre
            str1(2) = s.Value.Acredito

            Dim LVi1 As New ListViewItem(str1)
            'le meto el color si ya ingreso la persona en cuestion
            LVi1.BackColor = If(s.Value.Acredito > 0, Color.LightGreen, Nothing)
            ListViewRes.Items.Add(LVi1)
        Next
        ListViewRes.Columns(1).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)


        'act indicadores
        ToolStripProgressBar1.Maximum = ListView1.Items.Count
        ActulizarIndicadores()

    End Sub

    Private Sub AutoBusqueda(ByVal ListViewRes As ListView, ByVal TextBoxBusq As TextBox)
        ' AUTO BUSQUEDA BUSCA EL STRING EN EL TEXTO, LLENA EL LISTVIEW SEGUN LA COLECCION ACOMP2, VA EN EL KEYUP P KPRESS DEL TEXTBOXBUSQ
        Dim str1(2) As String
        ListViewRes.Items.Clear()

        If (TextBoxBusq.Text.Length = 0) Then
            CargarListaSinBuscar(ListViewRes)

            'HideResultado() 'oculto el list
            Return
        End If

        If (TextBoxBusq.Text.Length > 1) Then 'si escribe mas de 1 caracteres empieza a buscar
            For Each s As KeyValuePair(Of Int32, ColAComp3) In AcomP3 ' busco s en la coleccion acomp2 , s tiene dos componentes value y key, key bindea al id y valua nombre
                If s.Value.Nombre().Contains(TextBoxBusq.Text) Or s.Key.ToString.Contains(TextBoxBusq.Text) Then
                    '  Console.WriteLine("hallo:" + s.Key.ToString + s.Value)

                    'para llenar el listview
                    str1(0) = s.Key.ToString
                    str1(1) = s.Value.Nombre
                    str1(2) = s.Value.Acredito
                    Dim LVi1 As New ListViewItem(str1)

                    'le meto el color si ya ingreso la persona en cuestion
                    LVi1.BackColor = If(s.Value.Acredito > 0, Color.LightGreen, Nothing)

                    ListViewRes.Items.Add(LVi1)
                    ListViewRes.Columns(1).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)

                End If
            Next

            'si solo encontro uno automaticamente llego el textbox con ese valor
            If ListViewRes.Items.Count = 1 Then
                TextBoxBusq.Text = str1(1)
                TextBoxBusq.SelectAll()
                CargarFicha(str1(0))

                'If CInt(str1(2)) > 0 Then Return ' si ya se imprimio No preguntar nada
                'If MsgBox("Acreditar Reserva: " & str1(0) & " - " & str1(1), MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then
                '    TextBoxBusq.Text = ""
                '    Return
                'Else
                '    AcreditarReserva(str1(0))
                '    Return
                'End If

                'MsgBox("pp")
                Return
                'otros controles a llenar cuando encontro de una
                '                Label3.Text = str1(0)

            End If
            '''''

        End If

    End Sub

    Private Sub AutoBusquedaOpc(ByVal LlenarListSinBuscar As Boolean, ByVal ListViewRes As ListView, ByVal TextBoxBusq As TextBox)
        ' AUTO BUSQUEDA BUSCA EL STRING EN EL TEXTO, LLENA EL LISTVIEW SEGUN LA COLECCION ACOMP2, VA EN EL KEYUP P KPRESS DEL TEXTBOXBUSQ
        ListViewRes.Items.Clear()
        Dim str1(2) As String 'array para cada col de la list
        Select Case LlenarListSinBuscar
            Case True 'solo lleno la lista no busco, recorro la col

            Case False


                If (TextBoxBusq.Text.Length > 1) Then 'si escribe mas de 1 caracteres empieza a buscar

                    If (TextBoxBusq.Text.Length = 0) Then
                        'HideResultado() 'oculto el list
                        Return
                    End If
                    For Each s As KeyValuePair(Of Int32, ColAComp3) In AcomP3 ' busco s en la coleccion acomp2 , s tiene dos componentes value y key, key bindea al id y valua nombre
                        If s.Value.Nombre().Contains(TextBoxBusq.Text) Or s.Key.ToString.Contains(TextBoxBusq.Text) Then
                            '  Console.WriteLine("hallo:" + s.Key.ToString + s.Value)

                            'para llenar el listview
                            str1(0) = s.Key.ToString
                            str1(1) = s.Value.Nombre
                            str1(2) = s.Value.Ingreso
                            Dim LVi1 As New ListViewItem(str1)

                            'le meto el color si ya ingreso la persona en cuestion
                            LVi1.BackColor = If(s.Value.Ingreso = True, Color.LightGreen, Nothing)

                            ListViewRes.Items.Add(LVi1)
                            ListViewRes.Columns(1).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)

                        End If
                    Next

                    'si solo encontro uno automaticamente llego el textbox con ese valor
                    If ListViewRes.Items.Count = 1 Then
                        TextBoxBusq.Text = str1(1)
                        TextBoxBusq.SelectAll()


                        If MsgBox("Imprimir Reserva" & str1(0) & "-" & str1(1), MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then
                            TextBoxBusq.Text = ""
                        End If

                        'MsgBox("pp")
                        Return
                        'otros controles a llenar cuando encontro de una
                        '                Label3.Text = str1(0)

                    End If
                    '''''

                End If

        End Select


    End Sub

    Private Sub AutoBusquedaAcomp2(ByVal ListViewRes As ListView, ByVal TextBoxBusq As TextBox)
        ' AUTO BUSQUEDA BUSCA EL STRING EN EL TEXTO, LLENA EL LISTVIEW SEGUN LA COLECCION ACOMP2, VA EN EL KEYUP P KPRESS DEL TEXTBOXBUSQ
        ListViewRes.Items.Clear()
        Dim str1(1) As String

        If (TextBoxBusq.Text.Length > 1) Then 'si escribe mas de 1 caracteres empieza a buscar

            If (TextBoxBusq.Text.Length = 0) Then
                'HideResultado() 'oculto el list
                Return
            End If
            For Each s As KeyValuePair(Of Integer, String) In AComP2 ' busco s en la coleccion acomp2 , s tiene dos componentes value y key, key bindea al id y valua nombre
                If s.Value().Contains(TextBoxBusq.Text) Or s.Key.ToString.Contains(TextBoxBusq.Text) Then
                    '  Console.WriteLine("hallo:" + s.Key.ToString + s.Value)

                    'para llenar el listview
                    str1(0) = s.Key.ToString
                    str1(1) = s.Value

                    Dim LVi1 As New ListViewItem(str1)
                    ListViewRes.Items.Add(LVi1)

                End If
            Next

            'si solo encontro uno automaticamente llego el textbox con ese valor
            If ListViewRes.Items.Count = 1 Then
                TextBoxBusq.Text = str1(1)
                TextBoxBusq.SelectAll()


                If MsgBox("Imprimir Reserva" & str1(0) & "-" & str1(1), MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then
                    TextBoxBusq.Text = ""
                End If

                'MsgBox("pp")
                Return
                'otros controles a llenar cuando encontro de una
                '                Label3.Text = str1(0)

            End If
            '''''

        End If

    End Sub

    Private Sub ActulizarIndicadores()
        '       ToolStripProgressBar1.Value = CantIngresados
        '      ToolStripLabel1.Text = CantIngresados & " de " & ToolStripProgressBar1.Maximum

        ToolStripProgressBar1.Maximum = TotalReservas
        ToolStripProgressBar1.Value = CantResAcreditadas
        ToolStripLabel1.Text = CantResAcreditadas & " de " & ToolStripProgressBar1.Maximum
        ToolStripProgressBar1.ToolTipText = "Faltan " & ToolStripProgressBar1.Maximum - ToolStripProgressBar1.Value
        '        ToolStripProgressBar1.Visible = False

        ToolStripProgressBar2.Maximum = CantPersAcreditadas
        ToolStripProgressBar2.Value = CantPersIngresadas
        ToolStripLabel2.Text = CantPersIngresadas & " de " & ToolStripProgressBar2.Maximum
        ToolStripProgressBar2.ToolTipText = "Faltan " & ToolStripProgressBar2.Maximum - ToolStripProgressBar2.Value
        ToolStripProgressBar2.Visible = False



        ToolStripProgressBar3.Maximum = TotalPers
        ToolStripProgressBar3.Value = CantPersAcreditadas
        ToolStripLabel4.Text = CantPersAcreditadas & " de " & ToolStripProgressBar3.Maximum
        ToolStripProgressBar3.ToolTipText = "Faltan " & ToolStripProgressBar3.Maximum - ToolStripProgressBar3.Value
        ToolStripProgressBar3.Visible = False

        ToolStripProgressBar4.Maximum = TotalReservas
        ToolStripProgressBar4.Value = CantResPagadas
        ToolStripProgressBar4.ToolTipText = "Faltan " & ToolStripProgressBar4.Maximum - ToolStripProgressBar4.Value
        TslCantPagadas.Text = CantResPagadas & " de " & TotalReservas

    End Sub

    Public Sub CargarComparador()
        AcomP3.Clear()
        CantIngresados = 0
        CantPersIngresadas = 0
        CantPersAcreditadas = 0
        TotalPers = 0
        CantResAcreditadas = 0
        CantResPagadas = 0
        '        lleno la coleccion comparadora, recorro la DT cual tiene toda la tabla con el select x fecha y le saco los 2 valores id y nombre y los meto en la coleecion de 2d, es la que voy a usar para buscar leugo
        For Each fila As DataRow In oReservas.DTReserva.Rows

            'comparador viejo de 2 valores, es mas rapido pero solo para 2 valores y no se puede agrandar
            '            AComP2.Add(fila.Item("IDReserva"), UCase(fila.Item("Nombre"))) 'meto key nro(debe ser unico sino aexplota) y valor string

            ' lleno el comparador nuevo v a contenr id nombre ingreso si o no , se puede agrandar.
            Dim acol As New ColAComp3
            acol.Nombre = UCase(fila.Item("Nombre"))
            '-----------
            acol.CantPax = fila.Item("cantpax")
            '------
            If Not IsDBNull(fila.Item("ImprimioToken")) Then
                acol.Acredito = fila.Item("ImprimioToken")
                If acol.Acredito > 0 Then
                    CantResAcreditadas = CantResAcreditadas + 1
                    CantPersAcreditadas = CantPersAcreditadas + acol.CantPax
                End If

            Else
                acol.Acredito = 0
            End If
            '------

            If Not IsDBNull(fila.Item("Ingreso")) Then
                acol.Ingreso = fila.Item("Ingreso")
                If acol.Ingreso = True Then
                    CantIngresados = CantIngresados + 1
                    CantPersIngresadas = CantPersIngresadas + acol.CantPax
                End If
            Else
                acol.Ingreso = False
            End If

            '------SI ABONO O NO
            If Not IsDBNull(fila.Item("pago")) Then
                acol.pago = fila.Item("pago")
                If acol.pago > 0 Then
                    CantResPagadas = CantResPagadas + 1

                End If

            Else
                acol.pago = 0
            End If

            '------SI ABONO O NO' SI PRECIO UNIT es cer SIGNIFICA QUE NO ABONA, POR ESO LO PONGO COMO PAGA
            If Not IsDBNull(fila.Item("preciounit")) Then
                acol.PrecioUnit = fila.Item("preciounit")
                If acol.PrecioUnit = 0 Then
                    '    CantResPagadas = CantResPagadas + 1

                End If

            Else
                acol.PrecioUnit = 0
            End If

            '------ABONA SEGUN TIPO DE fact
            If Not IsDBNull(fila.Item("tipofactura")) Then
                Dim vTipoFac As String = fila.Item("tipofactura")
                For Each s As KeyValuePair(Of Integer, String) In ColFacNoAbona ' busco s en la coleccion acomp2 , s tiene dos componentes value y key, key bindea al id y valua nombre
                    If s.Value().Contains(vTipoFac) Then
                        acol.abona = True
                    End If
                Next


            End If
            '----------
            
            '------


            'LLENO LA COL CON ID, CLASECONTODOS Las props
            AcomP3.Add(fila.Item("IDReserva"), acol)


            TotalPers = TotalPers + acol.CantPax

            'acol.Nombre = UCase(fila.Item("Nombre") acol.Ingreso =  fila.Item("ingreso")


        Next

        TotalReservas = AcomP3.Count

        'CODE VEJO PARA USAR CON EL EXCEL/*
        '        Dim Conf As New Datos.Configuracion 'llamo los datos de xml

        '        Dim BD1 As New Datos.BaseXLS
        '        BD1.Conectar(Conf.DevolverValor("ExcelPath"))
        '        Dim TB1 As New DataTable

        '        TB1 = BD1.TomarDatos("select * from [U-Res-Consulta X EStado Diario $] where [Nº Rva] is not null")
        '        For Each rw As DataRow In TB1.Rows
        '            AComP2.Add(rw.Item("Nº Rva"), UCase(rw.Item("Nombre de Reserva"))) 'meto key nro(debe ser unico sino aexplota) y valor string
        '        Next
        '        TB1.Dispose()
        'CODE VEJO PARA USAR CON EL EXCEL
    End Sub
    Public Sub CargarMoneda()
        For Each fila As DataRow In oSetup.Moneda.Rows

            If Not IsDBNull(fila.Item("idmoneda")) Then
                ColMonedaTC.Add(fila.Item("idmoneda"), fila.Item("valor"))
                ColMonedaDESC.Add(fila.Item("idmoneda"), fila.Item("descuento"))
            End If
        Next


        ComboBox8.DisplayMember = "key"
        ComboBox8.ValueMember = "value"
        ComboBox8.DataSource = New BindingSource(ColMonedaTC, Nothing)

    End Sub

    Public Sub CargarTarjetaCol()

        For Each fila As DataRow In oSetup.TCredito.Rows
            If Not IsDBNull(fila.Item("tarjeta")) Then
                ColTarjetaDESC.Add(fila.Item("tarjeta"), fila.Item("descuento"))
            End If
        Next
    End Sub
    Public Sub CargarTipoResCli()
        For Each fila As DataRow In oSetup.TipoReserva.Rows
            If Not IsDBNull(fila.Item("tipo")) Then
                ColTipoRes.Add(fila.Item("tipo"), fila.Item("tipo"))
            End If

        Next

        ComboBox6TipoCli.DisplayMember = "key"
        ComboBox6TipoCli.ValueMember = "value"
        ComboBox6TipoCli.DataSource = New BindingSource(ColTipoRes, Nothing)

        'ComboBox6TipoCli.DataSource = oSetup.TipoReserva
        'ComboBox6TipoCli.DisplayMember = "tipo"
        ComboBox6TipoCli.AutoCompleteSource = AutoCompleteSource.ListItems


    End Sub
    Public Sub CargarTipoFactura()
        For Each fila As DataRow In oSetup.TipoFactura.Rows
            If Not IsDBNull(fila.Item("factura")) Then
                ColTipoFactura.Add(fila.Item("factura"), fila.Item("factura"))
            End If

        Next

        ComboBox5TipoFactura.DisplayMember = "key"
        ComboBox5TipoFactura.ValueMember = "value"
        ComboBox5TipoFactura.DataSource = New BindingSource(ColTipoFactura, Nothing)


        'ComboBox5TipoFactura.DataSource = oSetup.TipoFactura()
        'ComboBox5TipoFactura.ValueMember = "Factura"
        'ComboBox8.DisplayMember = "Factura"
        ComboBox5TipoFactura.AutoCompleteSource = AutoCompleteSource.ListItems




    End Sub
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Thread.CurrentThread.CurrentCulture = New CultureInfo("es-AR")
        'Thread.CurrentThread.CurrentUICulture = New CultureInfo("es-AR")
        '--------------------------------------------------------------
        InitCultura()
        'EAF LISTBOX1 INIT
        ListView1.View = View.Details
        ListView1.Columns.Add("ID", 100)
        ListView1.Columns.Add("Nombre", 1000)
        ListView1.Columns.Add("pago", 0)
        ListView1.FullRowSelect = True
        ListView1.HeaderStyle = ColumnHeaderStyle.None


        ' Add any initialization after the InitializeComponent() call.

        'LLAMO LAS RESERVAS segun la fecha de, OJO debe haber reservas cargadas para la fecha que la llamo 
        ' Dim FechaReservas As New Date() ' deberia setear siempre el dia de Hoy, cuando ocurre la reservao

        'FECHA Y ABRO RESERVAS
        oSetup = New BIZL.SetUp
        FechaReservas = oSetup.FechaCena.Date
        MsgBox("Fecha de las reservas:" & FechaReservas & vbCrLf & "Si fuera erronea comunicarse con Administracion.")
        Me.Text = Me.Text & " - Fecha: " & FechaReservas
        oReservas = New BIZL.Reservas(FechaReservas)


        'EAF LLENAR AUTOCOMPLETE 

        CargarColNoabonan()
        CargarComparador()
        CargarListaSinBuscar(ListView1)
        'CAMBIAR COLOR DE LOS LABELS
        'CambiarLabelColorFG(Me.Controls, Color.Cyan)
        CargarComboServicios()
        ComboBox2TV.Items.Add("SI")
        ComboBox2TV.Items.Add("NO")
        ComboBox3TI.Items.Add("SI")
        ComboBox3TI.Items.Add("NO")
        ComboBox7Ctango.Items.Add("SI")
        ComboBox7Ctango.Items.Add("NO")

        'cargo combo tipo reserva o tipo de cli
        CargarTipoResCli()






        TextBox2TotalARS1.Text = 0 'sino falla el acalcular cuando quiere hace el calculo mientras carga el combomoneda
        'Label36TotalARS.Text = 0

        'ComboMoneda.Items.Clear()
        'ComboMoneda.DataSource = oSetup.Moneda
        'ComboMoneda.ValueMember = "IDMONEDA"
        'ComboMoneda.SelectedIndex = 1
        CargarTipoFactura()
        CargarMoneda()
        CargarTarjetaCol()


        ' DesHabilitarModificacion()
        ListView1.Items(0).Selected = True 'dejo el 1ero asi nunca queda vacio
        'CARGO COMBO TRANSFER
        'ComboBox2.SelectedIndex = 0

        TextBox1.Clear()
        TextBox1.Focus()

        oCobros = New BIZL.Cobros(FechaReservas)

    End Sub

    Private Sub TextBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Click
        TextBox1.SelectAll()
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        TextBox1.SelectAll()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If (Char.IsLower(e.KeyChar)) Then 'pasa cada key a ucase
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If

    End Sub


    Private Sub TextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyUp
        AutoBusqueda(ListView1, TextBox1)
    End Sub

    Private Sub ListView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView1.KeyDown
        If e.KeyValue = Keys.Enter Then
                CobrarFrm.Show()
        End If
    End Sub

    Public Sub CargarFicha(ByVal aIDlista As String)
        'lleno los labels etc
        vCantPax = oReservas.Item(aIDlista).CantidadPax
        vPrecioUnit = oReservas.Item(aIDlista).PrecioUnitario
        vTipoFactura = oReservas.Item(aIDlista).TipoFactura
        vPrepago = oReservas.Item(aIDlista).Prepago
        vServicio = oReservas.Item(aIDlista).Servicio
        vMesa = oReservas.Item(aIDlista).Mesa
        vMozo = oReservas.Item(aIDlista).MozoID
        vTipoCli = oReservas.Item(aIDlista).TipoReserva
        vTI = If(oReservas.Item(aIDlista).TransferIda > 0, "SI", "NO")
        vTv = If(oReservas.Item(aIDlista).TransferVuelta > 0, "SI", "NO")
        vObs = oReservas.Item(aIDlista).Observaciones
        vCTango = If(oReservas.Item(aIDlista).ClaseTango = True, "SI", "")
        vPago = oReservas.Item(aIDlista).Pago
        vAcredito = oReservas.Item(aIDlista).ImprimioToken
        vIngreso = oReservas.Item(aIDlista).Ingreso

        MaskedTextBox1Cant.Text = vCantPax
        MaskedTextBox2Mesa.Text = vMesa
        ComboBox4Mozo.Text = vMozo
        ComboBox1Servicios.Text = vServicio
        ComboBox6TipoCli.Text = vTipoCli
        ComboBox2TV.Text = vTv
        ComboBox3TI.Text = vTI
        TextBox2Obs.Text = vObs
        ComboBox7Ctango.Text = vCTango


        Label1.Text = aIDlista
        Label6.Text = oReservas.Item(aIDlista).Mesa
        Label7.Text = vCantPax
        Label5.Text = oReservas.Item(aIDlista).NombreReserva
        Label9.Text = oReservas.Item(aIDlista).Observaciones
        Label4.Text = oReservas.Item(aIDlista).TipoReserva
        Label8.Text = oReservas.Item(aIDlista).Servicio
        Label10.Text = If(oReservas.Item(aIDlista).TransferVuelta > 0, "SI", "NO")



        'PRECIOS Y ACORBRAR
        TextBox4PrecioUnit.Text = vPrecioUnit
        TextBox3Prepago.Text = vPrepago
        '        TextBox2TotalARS.Text = vPrecioUnit * vCantPax - vpre
        'Label36TotalARS.Text = Math.Round((vPrecioUnit * vCantPax - vPrepago), 2) 'redondeo el nro final con 2 digitos
        TextBox2TotalARS1.Text = Math.Round((vPrecioUnit * vCantPax - vPrepago), 2)

        CalcularACobrarMoneda()

        'TextBox5ACobFinalARS.Text = TextBox2TotalARS.Text
        ComboBox5TipoFactura.Text = vTipoFactura

        If vPago > 0 Then
            Label21.Text = "ABONADO"
            Label21.Visible = True
        Else
            Label21.Visible = False
        End If

        If vPrecioUnit = 0 Then
            Label21.Text = "NO ABONA"
            Label21.Visible = True
        End If

        'ICONITOS DE ESTADO DE FICHA
        PictureBox1Acred.Visible = False
        PictureBox2Ingreso.Visible = False
        PictureBox3.Visible = False

        If vAcredito > 0 Then
            PictureBox1Acred.Visible = True
        End If
        If vIngreso = True Then
            PictureBox2Ingreso.Visible = True
        End If
        If vPago > 0 Then
            PictureBox3.Visible = True
        End If
        'FINESTADO

        'modificadores
        '    ComboBox1.Visible = False
        '        ComboBox1Servicios.Text = oReservas.Item(aIDlista).Servicio
        'MaskedTextBox1.Visible = False
        '       MaskedTextBox1Cant.Text = Label7.Text
    End Sub
    Private Sub CargarColNoabonan()
        ColFacNoAbona.Add(1, "Prepago")
        ColFacNoAbona.Add(2, "Cta.Cte.c/Voucher")
        ColFacNoAbona.Add(3, "Precompra")
        ColFacNoAbona.Add(4, "NO ABONA")
        ColFacNoAbona.Add(5, "Prepago")

    End Sub

    Private Sub CargarComboServicios()
        Dim TServicios As New BIZL.ServiciosCls
        ComboBox1Servicios.DataSource = TServicios.TiposServicio
        ComboBox1Servicios.ValueMember = "Servicio"
        ComboBox1Servicios.DisplayMember = "Servicio"
        ComboBox1Servicios.AutoCompleteSource = AutoCompleteSource.ListItems

    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        If ListView1.SelectedItems.Count > 0 Then
            '            TextBox1.Text = ListView1.SelectedItems(0).SubItems(1).Text 'meto lo que selecciono en el txtbox
            IDLista = ListView1.SelectedItems(0).SubItems(0).Text.ToString

            CargarFicha(IDLista)

        End If
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CerrarSis.Click

        If MsgBox("Salir del sistema?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Me.Dispose()
        End If

    End Sub

    Private Sub FiltroIngresados_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FiltroIngresados.Click
        Dim vtotars As Double = (oCobros.TotalUSD * ColMonedaTC("USD")) + (oCobros.TotalBRL * ColMonedaTC("BRL")) + (oCobros.TotalEUR * ColMonedaTC("EUR")) + oCobros.TotalTARJETA + oCobros.TotalARS


        MsgBox("ARS:" & oCobros.TotalARS & " BRL:" & oCobros.TotalBRL & " USD:" & oCobros.TotalUSD & " EUR:" & oCobros.TotalEUR & " TARJ:" & oCobros.TotalTARJETA & vbCrLf & "Total en ARS=" & vtotars)

    End Sub

    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel1.Click

    End Sub


    Private Sub AcreditarReserva()
        If Not ListView1.SelectedItems.Count = 0 Then 'si no hay nada seleccionado no hace nada


            If CInt(ListView1.SelectedItems(0).SubItems(2).Text) > 0 Then 'si ya acredito avisa
                If MsgBox("Ya se acredito x" & ListView1.SelectedItems(0).SubItems(2).Text & ", Desea imprimir acreditacion nuevamente?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Return
                End If
            End If
            oReservas.Item(IDLista).Acreditar()
            CargarComparador()
            ListView1.Items(ListView1.FocusedItem.Index).SubItems(2).Text = 1 + ListView1.Items(ListView1.FocusedItem.Index).SubItems(2).Text
            ListView1.Items(ListView1.FocusedItem.Index).BackColor = Color.LightGreen
            ActulizarIndicadores()
            ImprimirToken()
        End If
    End Sub
    Public Sub AcreditarReserva(ByVal aidreserva As Integer)
        oReservas.Item(aidreserva).Acreditar()
        CargarComparador()
        ActulizarIndicadores()
        'ImprimirToken()
    End Sub

    Public Sub ImprimirToken()
        Imprimir(True)

    End Sub

    Private Sub WalkinBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WalkinBtn.Click
        Walkin.Show()

    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        'cargo fuente
        Dim privateFonts As New System.Drawing.Text.PrivateFontCollection()
        privateFonts.AddFontFile("free3of9.ttf")
        Dim Font1 As New System.Drawing.Font(privateFonts.Families(0), 52) 'tam fuente code barra

        Dim Parrafo1 As String
        Parrafo1 = Label5.Text & vbCrLf & Label7.Text _
        & " - " & Label8.Text


        Dim imgLogo As Image
        imgLogo = Image.FromFile("MT_logo2byn.png")

        'dibujo toda la cosa
        e.Graphics.DrawImage(imgLogo, 10, 1)
        e.Graphics.DrawString(Parrafo1, New Font("Arial", 12, FontStyle.Bold), Brushes.Black, 5, 50)
        '	CODIGO DE BARRA->
        e.Graphics.DrawString(Label6.Text & Label1.Text, New Font("Arial", 13, FontStyle.Bold), Brushes.Black, 100, 100) 'chars barra
        e.Graphics.DrawString("*" & Label1.Text & "*", Font1, Brushes.Black, 15, 120) 'barra




        'Label1.Text = IDLista
        'Label6.Text = "Mesa:" & oReservas.Item(IDLista).Mesa
        'Label7.Text = "Cant:" & oReservas.Item(IDLista).CantidadPax
        'Label5.Text = oReservas.Item(IDLista).NombreReserva
        'Label9.Text = oReservas.Item(IDLista).Observaciones
        'Label4.Text = oReservas.Item(IDLista).TipoReserva
        'Label8.Text = oReservas.Item(IDLista).Servicio


    End Sub

    'Private Sub Imprimir()
    '    PrintDialog1.Document = PrintDocument1
    '    If (PrintDialog1.ShowDialog() = DialogResult.OK) Then
    '        PrintDocument1.Print()
    '    End If
    'End Sub
    Private Sub Imprimir(ByVal SinDialog As Boolean)

        If SinDialog = True Then
            '    PrintDialog1.Document = PrintDocument1
            '        If (PrintDialog1.ShowDialog() = DialogResult.OK) Then
            PrintDocument1.Print()
            'End If
        End If

        If SinDialog = False Then
            PrintDialog1.Document = PrintDocument1
            If (PrintDialog1.ShowDialog() = DialogResult.OK) Then
                PrintDocument1.Print()
            End If
        End If

    End Sub


    Private Sub ImprimirDeNuevo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImprimirDeNuevo.Click
        Imprimir(False)
    End Sub

    Private Sub btsModificaRes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btsCobrarRes.Click

        If vAcredito <= 0 Then

            MsgBox("La reserva seleccionada no ingreso, no se puede realizar el cobro")

        Else

            CobrarFrm.Show()
        End If

    End Sub

    Private Sub HabilitarModificacion()
        ComboBox1Servicios.Visible = True
        ComboBox2TV.Visible = True
        btModOK.Visible = True
        btModCancel.Visible = True
        MaskedTextBox1Cant.Visible = True

    End Sub

    Private Sub DesHabilitarModificacion()
        ComboBox1Servicios.Visible = False
        ComboBox2TV.Visible = False
        btModOK.Visible = False
        btModCancel.Visible = False
        MaskedTextBox1Cant.Visible = False

    End Sub

    Private Sub ModificarReserva(ByVal IDreserva As Integer)

        If vCantPax <> MaskedTextBox1Cant.Text Then
            'cambio cantidada
            '    MsgBox("cambio cantidad")
            oReservas.Item(IDreserva).CantidadPax = CInt(MaskedTextBox1Cant.Text)
            Label7.Text = MaskedTextBox1Cant.Text
        End If

        If ComboBox1Servicios.Text.Substring(2) <> vServicio.Substring(2) Then 'substring 2 , solo se fija si cambiaron las 2primeras letras
            'cambio servicio
            MsgBox("CAMBIO DE SERVICIO DETALLAR")
            CambioServicio.Show()
        End If

        If ComboBox2TV.Text <> vTv Then
            'cambio transfer
            oReservas.Item(IDreserva).TransferVuelta = If(ComboBox2TV.Text = "SI", 1, 0)
            Label10.Text = ComboBox2TV.Text
        End If

        If ComboBox6TipoCli.Text <> vTipoCli Then
            oReservas.Item(IDreserva).TipoReserva = ComboBox6TipoCli.Text
        End If

        If MaskedTextBox2Mesa.Text <> vMesa Then
            oReservas.Item(IDreserva).Mesa = MaskedTextBox2Mesa.Text
        End If

        If ComboBox3TI.Text <> vTI Then
            'cambio transfer
            oReservas.Item(IDreserva).TransferIda = If(ComboBox3TI.Text = "SI", 1, 0)
        End If

        If TextBox2Obs.Text <> vObs Then
            oReservas.Item(IDreserva).Observaciones = TextBox2Obs.Text
        End If

        If TextBox4PrecioUnit.Text <> vPrecioUnit Then
            oReservas.Item(IDreserva).PrecioUnitario = TextBox4PrecioUnit.Text
        End If

        If TextBox3Prepago.Text <> vPrepago Then
            oReservas.Item(IDreserva).Prepago = TextBox3Prepago.Text
        End If

        If ComboBox7Ctango.Text <> vCTango Then
            oReservas.Item(IDreserva).ClaseTango = If(ComboBox7Ctango.Text = "SI", 1, 0)
        End If

        If ComboBox5TipoFactura.Text <> vTipoFactura Then
            oReservas.Item(IDreserva).TipoFactura = ComboBox5TipoFactura.Text

        End If

        '        DesHabilitarModificacion()
        CargarComparador()
        CargarListaSinBuscar(ListView1)
        'CargarFicha(IDreserva)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ModificarReserva(CInt(Label1.Text))
    End Sub

    Private Sub btModCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        DesHabilitarModificacion()
    End Sub

    Private Sub MaskedTextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        MaskedTextBox1Cant.SelectAll()
    End Sub

    Private Sub MaskedTextBox1_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs)

    End Sub

    Private Sub MaskedTextBox2Mesa_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MaskedTextBox2Mesa.GotFocus
        MaskedTextBox2Mesa.SelectAll()
    End Sub

    Private Sub MaskedTextBox2Mesa_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MaskedTextBox2Mesa.KeyPress
        If (e.KeyChar = " ") Then
            e.KeyChar = ""
        End If
    End Sub

    Private Sub MaskedTextBox2_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs) Handles MaskedTextBox2Mesa.MaskInputRejected

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    'Private Sub ComboMoneda_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    '  Label29DescuentoMoneda.Text = "CAMBIO:" & ColMonedaTC(ComboMoneda.Text)
    '    Label29DescuentoMoneda.Text = "PAR:" & ColMonedaTC(ComboMoneda.Text) & " DESC:" & 100 * ColMonedaDESC(ComboMoneda.Text) & "%"


    'End Sub

    'Private Sub TextBox7ACobMoneda_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If IsNumeric(TextBox7ACobMoneda.Text) Then

    '        CalcularACobrar(TextBox8ACobTarj.Text, TextBox6ACobARS2.Text, TextBox7ACobMoneda.Text)

    '    End If
    'End Sub

    Private Sub TextBox5ACobARS_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    'Private Sub CalcularACobrar(ByVal CantTarjeta As Double, ByVal CantARS As Double, ByVal CantMonExt As Double)

    '    If IsNumeric(CantTarjeta) And IsNumeric(CantARS) And IsNumeric(CantMonExt) Then
    '        'devuelve lo que va a cobrar, incluyendo descuentos etc, en ARS
    '        Dim vTCambio As Double = ColMonedaTC(ComboBox8.Text)
    '        Dim vDescTotal As Double = CantMonExt - CantMonExt * ColMonedaDESC(ComboBox8.Text)
    '        '    Dim vMonExtenARS As Double = CantMonExt * vTCambio - vDescTotal
    '        Dim vMonExtenARS As Double = CantMonExt * vTCambio

    '        TextBox5ACobFinalARS.Text = CantTarjeta + CantARS + vMonExtenARS

    '    End If

    'End Sub

    Private Sub CalcularACobrarMoneda()

        Label34.Text = "PAR:" & ColMonedaTC(ComboBox8.Text) & " DESC:" & 100 * ColMonedaDESC(ComboBox8.Text) & "%"

        Dim vTCambio As Double = ColMonedaTC(ComboBox8.Text)
        Dim vCantMonExt As Double = TextBox2TotalARS1.Text / vTCambio
        Dim vMonExtConDesc As Double = vCantMonExt - vCantMonExt * ColMonedaDESC(ComboBox8.Text)

        TextBox2TotalARS.Text = Math.Round(vMonExtConDesc, 2)

    End Sub

    Private Sub Label21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox8.SelectedIndexChanged
        CalcularACobrarMoneda()

    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub MaskedTextBox1_GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles MaskedTextBox1Cant.GotFocus
        MaskedTextBox1Cant.SelectAll()

    End Sub


    Private Sub ComboBox6TipoCli_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox6TipoCli.SelectedIndexChanged
        ModificoAlgo(vTipoCli, ComboBox6TipoCli.Text)


    End Sub
    Private Function ModificoAlgo(ByVal inicial As String, ByVal final As String) As Boolean
        If inicial <> final Then
            cambioFicha = cambioFicha + 1
            btModOK.Visible = True
            btModCancel.Visible = True
            TestLabel.Text = cambioFicha
            'If cambioFicha > 1 Then
            '    cambioFicha = 0
            'End If
            Return True
        Else
            cambioFicha = cambioFicha - 1
            If (cambioFicha <= 0) Then 'pone en invisible si no hay modis previos
                btModOK.Visible = False
                btModCancel.Visible = False
                cambioFicha = 0
            Else


            End If

            TestLabel.Text = cambioFicha
            Return False
        End If


    End Function

    Private Sub ComboBox1Servicios_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1Servicios.SelectedIndexChanged
        ModificoAlgo(vServicio, ComboBox1Servicios.Text)
    End Sub

    Private Sub MaskedTextBox1Cant_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MaskedTextBox1Cant.KeyPress
        If (e.KeyChar = " ") Then
            e.KeyChar = ""
        End If
    End Sub

    Private Sub MaskedTextBox1Cant_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs) Handles MaskedTextBox1Cant.MaskInputRejected



    End Sub

    Private Sub MaskedTextBox1Cant_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MaskedTextBox1Cant.TextChanged
        ModificoAlgo(vCantPax, MaskedTextBox1Cant.Text)

    End Sub

    Private Sub MaskedTextBox2Mesa_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MaskedTextBox2Mesa.TextChanged
        ModificoAlgo(vMesa, MaskedTextBox2Mesa.Text)
    End Sub

    Private Sub ComboBox3TI_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3TI.SelectedIndexChanged
        ModificoAlgo(vTI, ComboBox3TI.Text)

    End Sub

    Private Sub ComboBox2TV_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2TV.SelectedIndexChanged
        ModificoAlgo(vTv, ComboBox2TV.Text)
    End Sub

    Private Sub ComboBox7Ctango_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox7Ctango.SelectedIndexChanged
        ModificoAlgo(vCTango, ComboBox7Ctango.Text)
    End Sub

    Private Sub TextBox2Obs_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2Obs.TextChanged
        ModificoAlgo(vObs, TextBox2Obs.Text)
    End Sub

    Private Sub ComboBox5TipoFactura_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5TipoFactura.SelectedIndexChanged
        ModificoAlgo(vTipoFactura, ComboBox5TipoFactura.Text)


    End Sub

    Private Sub TextBox4PrecioUnit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4PrecioUnit.KeyPress
        e.Handled = onlyNumbers(e.KeyChar)
    End Sub

    Private Sub TextBox4PrecioUnit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4PrecioUnit.TextChanged
        If Not IsNumeric(TextBox4PrecioUnit.Text) Then
            TextBox4PrecioUnit.Text = 0
            TextBox4PrecioUnit.SelectAll()
        End If
        ModificoAlgo(vPrecioUnit, TextBox4PrecioUnit.Text)

    End Sub

    Private Sub TextBox3Prepago_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3Prepago.KeyPress
        e.Handled = onlyNumbers(e.KeyChar)
    End Sub

    Private Sub TextBox3Prepago_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3Prepago.TextChanged
        ModificoAlgo(vPrepago, TextBox3Prepago.Text)
    End Sub

    Private Sub btModOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btModOK.Click
        ModificarReserva(IDLista)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Focus()

    End Sub
    Private Sub RefrescarTodo()
        oReservas.SincronizarBase()
        CargarComparador()
        CargarListaSinBuscar(ListView1)
        ActulizarIndicadores()
    End Sub

    Private Sub tsbRefrescarTodo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbRefrescarTodo.Click
        RefrescarTodo()
    End Sub
    Public Sub IngresarReserva(ByVal aidreserva As Integer)
        oReservas.Item(aidreserva).Ingresar()
        CargarComparador()
        CargarListaSinBuscar(ListView1)
        ActulizarIndicadores()
    End Sub

    Private Sub tsbtIngresar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbtIngresar.Click
        If MsgBox("Desea Ingresar " & IDLista & " ", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            AcreditarReserva(IDLista)
            IngresarReserva(IDLista)
            RefrescarTodo()
            CargarFicha(IDLista)
        End If
    End Sub

    Private Sub Label23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Public Function onlyNumbers(ByVal KeyChar As Char) As Boolean
        Dim allowedChars As String

        allowedChars = "0123456789."

        If allowedChars.IndexOf(KeyChar) = -1 And (Asc(KeyChar)) <> 8 Then
            Return True
        End If

        Return False
    End Function

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs)

        MsgBox("ARS:" & oCobros.TotalARS & " BRL:" & oCobros.TotalBRL & " USD:" & oCobros.TotalUSD & " EUR:" & oCobros.TotalEUR & " TARJ:" & oCobros.TotalTARJETA)
        Return
    End Sub
End Class
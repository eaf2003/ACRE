Imports System
Imports System.Threading
Imports System.Globalization


Public Class Acreditaciones
    Public Shared oReservas As BIZL.Reservas
    Public Shared oSetup As BIZL.SetUp
    Public IDLista
    Public Shared FechaReservas As Date

    Public CantResAcreditadas As Integer
    Public CantIngresados As Integer
    Public CantClases As Integer
    Public TotalReservas As Integer
    Public CantPersAcreditadas As Integer
    Public CantPersIngresadas As Integer
    Public TotalPers As Integer
    Private vServicio As String
    Private vClaseT As Boolean
    Private vTipoFactura As String

    Dim TServicios As New BIZL.ServiciosCls
    Dim oCfg As New Datos.Configuracion


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
                If CInt(str1(2)) > 0 Then Return ' si ya se imprimio No preguntar nada
                If MsgBox("Acreditar Reserva: " & str1(0) & " - " & str1(1), MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then
                    TextBoxBusq.Text = ""
                    Return
                Else
                    AcreditarReserva(str1(0))
                    Return
                End If

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


        ToolStripProgressBar2.Maximum = CantPersAcreditadas
        ToolStripProgressBar2.Value = CantPersIngresadas
        ToolStripLabel2.Text = CantPersIngresadas & " de " & ToolStripProgressBar2.Maximum
        ToolStripProgressBar2.ToolTipText = "Faltan " & ToolStripProgressBar2.Maximum - ToolStripProgressBar2.Value

        ToolStripProgressBar3.Maximum = TotalPers
        ToolStripProgressBar3.Value = CantPersAcreditadas
        ToolStripLabel4.Text = CantPersAcreditadas & " de " & ToolStripProgressBar3.Maximum
        ToolStripProgressBar3.ToolTipText = "Faltan " & ToolStripProgressBar3.Maximum - ToolStripProgressBar3.Value

        tsLclases.Text = "Clases: " & CantClases

    End Sub

    Public Sub CargarComparador()
        AcomP3.Clear()
        CantIngresados = 0
        CantPersIngresadas = 0
        CantPersAcreditadas = 0
        TotalPers = 0
        CantResAcreditadas = 0
        CantClases = 0
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
            '------

            If Not IsDBNull(fila.Item("clasetango")) Then
                acol.ClaseTango = fila.Item("clasetango")
                If acol.ClaseTango = True Then
                    CantClases = CantClases + 1
                End If
            Else
                acol.Ingreso = False
            End If





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

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        'por cosas de fechas y demas seteo la cultura fija
        InitCultura()
        '  ShowCurrentCulture()

        '        Thread.CurrentThread.CurrentCulture.DateTimeFormat = dd / MM / yyyy

        'EAF LISTBOX1 INIT
        ListView1.View = View.Details
        ListView1.Columns.Add("ID", 100)
        ListView1.Columns.Add("Nombre", 1000)
        ListView1.Columns.Add("Acredito", 0)
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


        CargarComparador()
        CargarListaSinBuscar(ListView1)
        'CAMBIAR COLOR DE LOS LABELS
        'CambiarLabelColorFG(Me.Controls, Color.Cyan)
        CargarComboServicios()
        CargarServiciosDescripcion()

        ComboBox2.Items.Add("SI")
        ComboBox2.Items.Add("NO")

        ComboBox3ClaseT.Items.Add("SI")
        ComboBox3ClaseT.Items.Add("NO")


        CargarTipoResCli()
        CargarTipoFactura()

        DesHabilitarModificacion()
        ListView1.Items(0).Selected = True 'dejo el 1ero asi nunca queda vacio
        'CARGO COMBO TRANSFER
        'ComboBox2.SelectedIndex = 0

        TextBox1.Clear()
        TextBox1.Focus()

    End Sub
    Public Sub CargarTipoFactura()
        For Each fila As DataRow In oSetup.TipoFactura.Rows
            If Not IsDBNull(fila.Item("factura")) Then
                ColTipoFactura.Add(fila.Item("factura"), fila.Item("factura"))
            End If

        Next


    End Sub

    Public Sub CargarServiciosDescripcion()

        For Each fila As DataRow In TServicios.TiposServicio.Rows
            If Not IsDBNull(fila.Item("descripcion")) Then
                ColServicioDescripcion.Add(fila.Item("servicio"), fila.Item("descripcion"))
            End If

        Next

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
            AcreditarReserva()
        End If
    End Sub

    Public Sub CargarFicha(ByVal aIDlista As String)
        vServicio = oReservas.Item(aIDlista).Servicio
        vClaseT = oReservas.Item(aIDlista).ClaseTango
        vTipoFactura = oReservas.Item(aIDlista).TipoFactura
        'lleno los labels etc
        Label1.Text = aIDlista
        Label6.Text = oReservas.Item(aIDlista).Mesa
        Label7.Text = oReservas.Item(aIDlista).CantidadPax
        Label5.Text = oReservas.Item(aIDlista).NombreReserva
        Label9.Text = oReservas.Item(aIDlista).Observaciones
        Label4.Text = oReservas.Item(aIDlista).TipoReserva
        Label8.Text = vServicio
        Label17.Text = vTipoFactura
        Label10.Text = If(oReservas.Item(aIDlista).TransferVuelta > 0, "SI", "NO")
        Label17claseT.Text = If(vClaseT = True, "SI", "NO")
        ComboBox3ClaseT.Text = Label17claseT.Text
        'modificadores
        ComboBox1.Visible = False
        ComboBox1.Text = oReservas.Item(aIDlista).Servicio
        ComboBox2.Text = Label10.Text
        MaskedTextBox1.Visible = False
        MaskedTextBox1.Text = Label7.Text
    End Sub
    Private Sub CargarComboServicios()

        ComboBox1.DataSource = TServicios.TiposServicio()
        ComboBox1.ValueMember = "Servicio"
        ComboBox1.DisplayMember = "Servicio"
        ComboBox1.AutoCompleteSource = AutoCompleteSource.ListItems

    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        If ListView1.SelectedItems.Count > 0 Then
            TextBox1.Text = ListView1.SelectedItems(0).SubItems(1).Text 'meto lo que selecciono en el txtbox
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
        If FiltroIngresados.CheckState = CheckState.Checked Then

        Else

        End If

    End Sub

    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel1.Click

    End Sub

    Private Sub Acreditar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Acreditar.Click
        AcreditarReserva()
    End Sub

    Private Sub AcreditarReserva()
        If Not ListView1.SelectedItems.Count = 0 Then 'si no hay nada seleccionado no hace nada

            If CInt(ListView1.SelectedItems(0).SubItems(2).Text) > 0 Then 'si ya acredito avisa
                If MsgBox("Ya se acredito x" & ListView1.SelectedItems(0).SubItems(2).Text & ", Desea imprimir acreditacion nuevamente?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Return
                End If
            End If
            QuiereClaseTango(IDLista)
            oReservas.Item(IDLista).Acreditar()
            CargarComparador()
            ListView1.Items(ListView1.FocusedItem.Index).SubItems(2).Text = 1 + ListView1.Items(ListView1.FocusedItem.Index).SubItems(2).Text
            ListView1.Items(ListView1.FocusedItem.Index).BackColor = Color.LightGreen
            ActulizarIndicadores()
            ImprimirToken()
        End If
    End Sub

    Public Sub AcreditarReserva(ByVal aidreserva As Integer)
        QuiereClaseTango(aidreserva)
        oReservas.Item(aidreserva).Acreditar()
        CargarComparador()
        ActulizarIndicadores()
        ImprimirToken()
    End Sub


    Public Sub QuiereClaseTango(ByVal aidreserva As Integer)
        If MsgBox("Quiere Clase de tango?", MsgBoxStyle.YesNo, "CLASE DE TANGO") = MsgBoxResult.Yes Then
            oReservas.Item(aidreserva).ClaseTango = True
        Else
            oReservas.Item(aidreserva).ClaseTango = False
        End If

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
        & " - " & Label8.Text & "-" & Label7.Text
        Dim Parrafo2 As String = If(ColServicioDescripcion.ContainsKey(vServicio), ColServicioDescripcion(vServicio), Nothing)


        Dim imgLogo As Image
        imgLogo = Image.FromFile("MT_logo2byn.png")

        'dibujo toda la cosa
        e.Graphics.DrawImage(imgLogo, 10, 1)
        e.Graphics.DrawString(Parrafo1, New Font("Arial", 12, FontStyle.Bold), Brushes.Black, 5, 50)
        e.Graphics.DrawString(Parrafo2, New Font("Arial", 10, FontStyle.Regular), Brushes.Black, 5, 85) 'descripcion de servicio
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


    Private Sub Imprimir(ByVal SinDialog As Boolean)
        Dim vPRN As String = oCfg.DevolverValor("PRN", "Acreditador.xml")
        Try


            If SinDialog = True Then
                '    PrintDialog1.Document = PrintDocument1
                '        If (PrintDialog1.ShowDialog() = DialogResult.OK) Then
                PrintDocument1.PrinterSettings.PrinterName = vPRN
                PrintDocument1.Print()
                'End If
            End If

            If SinDialog = False Then
                If (PrintDialog1.ShowDialog() = DialogResult.OK) Then
                    PrintDialog1.Document = PrintDocument1
                    PrintDocument1.Print()
                End If

            End If

        Catch ex As Exception
            MsgBox("Error en Acreditaciones.Imprimir PRN:" & vPRN & " ERROR:" & ex.Message)
        End Try


    End Sub


    Private Sub ImprimirDeNuevo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImprimirDeNuevo.Click
        Imprimir(False)
    End Sub

    Private Sub btsModificaRes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btsModificaRes.Click
        HabilitarModificacion()
    End Sub

    Private Sub HabilitarModificacion()
        ComboBox3ClaseT.Visible = True
        ComboBox1.Visible = True
        ComboBox2.Visible = True
        btModOK.Visible = True
        btModCancel.Visible = True
        MaskedTextBox1.Visible = True

    End Sub

    Private Sub DesHabilitarModificacion()
        ComboBox3ClaseT.Visible = False
        ComboBox1.Visible = False
        ComboBox2.Visible = False
        btModOK.Visible = False
        btModCancel.Visible = False
        MaskedTextBox1.Visible = False

    End Sub

    Private Sub ModificarReserva(ByVal IDreserva As Integer)

        If Label7.Text <> MaskedTextBox1.Text Then
            'cambio cantidada
            '    MsgBox("cambio cantidad")
            oReservas.Item(IDreserva).CantidadPax = CInt(MaskedTextBox1.Text)
            Label7.Text = MaskedTextBox1.Text
        End If

        If ComboBox1.Text.Substring(2) <> Label8.Text.Substring(2) Then
            'cambio servicio
            MsgBox("CAMBIO DE SERVICIO DETALLAR")
            CambioServicio.Show()
        End If


        If ComboBox2.Text <> Label10.Text Then
            'cambio transfer
            oReservas.Item(IDreserva).TransferVuelta = If(ComboBox2.Text = "SI", 1, 0)
            Label10.Text = ComboBox2.Text
        End If


        If ComboBox3ClaseT.Text <> Label17claseT.Text Then
            'cambio transfer
            oReservas.Item(IDreserva).ClaseTango = If(ComboBox3ClaseT.Text = "SI", True, False)
            Label17claseT.Text = ComboBox3ClaseT.Text
        End If

        DesHabilitarModificacion()
        CargarComparador()
        CargarListaSinBuscar(ListView1)
        'CargarFicha(IDreserva)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btModOK.Click
        ModificarReserva(CInt(Label1.Text))
    End Sub

    Private Sub btModCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btModCancel.Click
        DesHabilitarModificacion()
    End Sub

    Private Sub MaskedTextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MaskedTextBox1.GotFocus
        MaskedTextBox1.SelectAll()
    End Sub

    Private Sub MaskedTextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MaskedTextBox1.KeyPress
        If (e.KeyChar = " ") Then
            e.KeyChar = ""
        End If
    End Sub

    Private Sub MaskedTextBox1_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs) Handles MaskedTextBox1.MaskInputRejected

    End Sub

    Public Sub CargarTipoResCli()
        For Each fila As DataRow In oSetup.TipoReserva.Rows
            If Not IsDBNull(fila.Item("tipo")) Then
                ColTipoRes.Add(fila.Item("tipo"), fila.Item("tipo"))
            End If
        Next
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        RefrescarTodo()
    End Sub
    Private Sub RefrescarTodo()
        oReservas.SincronizarBase()
        CargarComparador()
        CargarListaSinBuscar(ListView1)
        ActulizarIndicadores()
    End Sub
End Class
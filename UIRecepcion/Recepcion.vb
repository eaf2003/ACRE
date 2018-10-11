Imports System.Threading
Imports System.Globalization
Public Class Recepcion
    Public oReservas As BIZL.Reservas
    Public Shared oSetup As BIZL.SetUp
    Public Shared FechaReservas As Date
    Public IDLista
    Public CantIngresados As Integer
    Public TotalReservas As Integer
    Public CantPersIngresadas As Integer
    Public TotalPers As Integer
    Private Sub CambiarLabelColorFG(ByVal controls As ControlCollection, ByVal fgColor As Color)
        If controls Is Nothing Then Return
        For Each C As Control In controls
            If TypeOf C Is Label Then DirectCast(C, Label).ForeColor = fgColor
            If C.HasChildren Then CambiarLabelColorFG(C.Controls, fgColor)
        Next
    End Sub

    Public Sub CargarListaSinBuscar(ByVal ListViewRes As ListView)
        ListViewRes.Items.Clear() 'fijarse de darle clear a los items y no a toda la lista, sino borro los headers tambien
        Dim str1(3) As String
        For Each s As KeyValuePair(Of Int32, ColAComp3) In AcomP3 ' busco s en la coleccion acomp2 , s tiene dos componentes value y key, key bindea al id y valua nombre
            'para llenar el listview
            str1(0) = s.Key.ToString
            str1(1) = s.Value.Nombre
            str1(2) = s.Value.Ingreso
            str1(3) = s.Value.CantPax
            Dim LVi1 As New ListViewItem(str1)
            'le meto el color si ya ingreso la persona en cuestion
            LVi1.BackColor = If(s.Value.Ingreso = True, Color.LightGreen, Nothing)
            ListViewRes.Items.Add(LVi1)
        Next
        ListViewRes.Columns(1).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent) 'resize contra el nombre


        'act indicadores
        ToolStripProgressBar1.Maximum = ListView1.Items.Count
        ActulizarIndicadores()

    End Sub
    Private Sub AutoBusquedaAcreditador(ByVal ListViewRes As ListView, ByVal TextBoxBusq As TextBox)
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

                If CInt(str1(2)) > 0 Then Return ' si ya se imprimio No preguntar nada
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
    Private Sub AutoBusqueda(ByVal ListViewRes As ListView, ByVal TextBoxBusq As TextBox)
        ' AUTO BUSQUEDA BUSCA EL STRING EN EL TEXTO, LLENA EL LISTVIEW SEGUN LA COLECCION ACOMP2, VA EN EL KEYUP P KPRESS DEL TEXTBOXBUSQ

        Dim str1(3) As String
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
                    str1(2) = s.Value.Ingreso
                    str1(3) = s.Value.CantPax
                    Dim LVi1 As New ListViewItem(str1)

                    'le meto el color si ya ingreso la persona en cuestion
                    LVi1.BackColor = If(s.Value.Ingreso = True, Color.LightGreen, Nothing)

                    ListViewRes.Items.Add(LVi1)

                End If
            Next

            'si solo encontro uno automaticamente llego el textbox con ese valor
            If ListViewRes.Items.Count = 1 Then
                ' TextBoxBusq.Text = str1(1)
                '  TextBoxBusq.SelectAll()
                IDLista = str1(0) 'idlista es el id que encontro
                CargarFicha(IDLista)
                '                MsgBox("INGRESANDO...")
                Ingresando.Show()
                IngresarReserva(IDLista)
                ' oReservas.Item(IDLista).Ingresar()

                '                CargarComparador()
                '               ActulizarIndicadores()
                TextBoxBusq.Clear()
                '              CargarListaSinBuscar(ListView1)
                'MsgBox("pp")
                Return
                'otros controles a llenar cuando encontro de una
                '                Label3.Text = str1(0)

            End If
            '''''
            ListViewRes.Columns(1).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)

        End If

    End Sub
    Private Sub AutoBusqueda(ByVal TextBoxBusq As TextBox)
        ' AUTO BUSQUEDA BUSCA EL STRING EN EL TEXTO, LLENA EL LISTVIEW SEGUN LA COLECCION ACOMP2, VA EN EL KEYUP P KPRESS DEL TEXTBOXBUSQ
        Dim str1(2) As String


        If (TextBoxBusq.Text.Length = 0) Then


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



                End If
            Next

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
        ToolStripProgressBar1.Maximum = TotalReservas
        ToolStripProgressBar1.Value = CantIngresados
        ToolStripLabel1.Text = CantIngresados & " de " & ToolStripProgressBar1.Maximum
        ToolStripProgressBar1.ToolTipText = "Faltan " & ToolStripProgressBar1.Maximum - ToolStripProgressBar1.Value

        ToolStripProgressBar2.Maximum = TotalPers
        ToolStripProgressBar2.Value = CantPersIngresadas
        ToolStripLabel2.Text = CantPersIngresadas & " de " & ToolStripProgressBar2.Maximum
        ToolStripProgressBar2.ToolTipText = "Faltan " & ToolStripProgressBar2.Maximum - ToolStripProgressBar2.Value
    End Sub

    Public Sub CargarComparador()
        AcomP3.Clear()
        CantIngresados = 0
        CantPersIngresadas = 0
        TotalPers = 0

        '        lleno la coleccion comparadora, recorro la DT cual tiene toda la tabla con el select x fecha y le saco los 2 valores id y nombre y los meto en la coleecion de 2d, es la que voy a usar para buscar leugo
        For Each fila As DataRow In oReservas.DTReserva.Rows
            ' lleno el comparador nuevo v a contenr id nombre ingreso si o no , se puede agrandar.
            Dim acol As New ColAComp3
            acol.Nombre = UCase(fila.Item("Nombre"))
            acol.CantPax = fila.Item("cantpax")

            If Not IsDBNull(fila.Item("Ingreso")) Then
                acol.Ingreso = fila.Item("Ingreso")
                If acol.Ingreso = True Then
                    CantIngresados = CantIngresados + 1
                    CantPersIngresadas = CantPersIngresadas + acol.CantPax
                End If
            Else
                acol.Ingreso = False
            End If

            AcomP3.Add(fila.Item("IDReserva"), acol) ' agrega = (ID# ,nobre,igreso) lleno el comparador

            'calculo la gente
            TotalPers = TotalPers + acol.CantPax

            'sigo con la proxima row

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
        InitCultura()
        '  Thread.CurrentThread.CurrentCulture = New CultureInfo("es-AR")
        ' Thread.CurrentThread.CurrentUICulture = New CultureInfo("es-AR")
        '--------------------------------------------------------------
        'EAF LISTBOX1 INIT
        ListView1.View = View.Details
        ListView1.Columns.Add("ID", 100)
        ListView1.Columns.Add("Nombre", 1000)
        ListView1.Columns.Add("Ingreso", 0)
        ListView1.Columns.Add("Cantidad", 60)
        ListView1.FullRowSelect = True
        ListView1.HeaderStyle = ColumnHeaderStyle.None
        '        ListView1.
        'LLAMO LAS RESERVAS segun la fecha de, OJO debe haber reservas cargadas para la fecha que la llamo 

        oSetup = New BIZL.SetUp
        FechaReservas = oSetup.FechaCena.Date

        'MsgBox("Fecha de las reserva' s:" & FechaReservas & vbCrLf & "Si fuera erronea comunicarse con Administracion.")

        '''' Dim FechaReservas As New Date(2014, 11, 7) ' deberia setear siempre el dia de Hoy, cuando ocurre la reserva
        '        oReservas = New BIZL.Reservas(FechaReservas, "ImprimioToken > 0", "CASE WHEN Ingreso IS NULL THEN 1 ELSE 0 END, Ingreso, CASE WHEN fechaimp IS NULL THEN 1980 END")

        oReservas = New BIZL.Reservas(FechaReservas, "ImprimioToken > 0", "CASE WHEN Ingreso IS NULL THEN 0 ELSE 1 END, Ingreso DESC, FechaImp")
        'si ingreso es null que los ponga abajo
        'EAF LLENAR AUTOCOMPLETE 

        CargarComparador()
        CargarListaSinBuscar(ListView1)
        'CAMBIAR COLOR DE LOS LABELS
        'CambiarLabelColorFG(Me.Controls, Color.Cyan)
        TextBox1.Focus()
        ActulizarIndicadores()

        Label1.Text = ""
        Label4.Text = ""
        Label5.Text = ""
        Label6.Text = ""
        Label7.Text = ""
        Label8.Text = ""

        Me.Text = Me.Text & " - Fecha: " & FechaReservas.Date

        Timer1.Interval = 8000 '8seconds
        Timer1.Start()

    End Sub

    Private Sub TextBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Click
        TextBox1.SelectAll()
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        TextBox1.SelectAll()
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyValue = Keys.Enter Then
            RefrescarTodo() 'sincronizo a la base y recargo por si hubo un acreditado apenas antes de buscar
            AutoBusqueda(ListView1, TextBox1)
            TextBox1.Clear()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If (Char.IsLower(e.KeyChar)) Then 'pasa cada key a ucase
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If

    End Sub


    Private Sub TextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyUp
        'AutoBusqueda(ListView1, TextBox1)
    End Sub

    Private Sub ListView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView1.KeyDown
        If e.KeyValue = Keys.Enter Then
            IngresarReserva()
        End If
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        If ListView1.SelectedItems.Count > 0 Then
            '  TextBox1.Text = ListView1.SelectedItems(0).SubItems(1).Text 'meto lo que selecciono en el txtbox
            IDLista = ListView1.SelectedItems(0).SubItems(0).Text.ToString

            CargarFicha(IDLista)

            'LLENO LOS LABELS

            'Label1.Text = IDLista
            'Label6.Text = oReservas.Item(IDLista).Mesa
            'Label7.Text = "Cant:" & oReservas.Item(IDLista).CantidadPax
            'Label5.Text = oReservas.Item(IDLista).NombreReserva
            'Label9.Text = oReservas.Item(IDLista).Observaciones
            'Label4.Text = oReservas.Item(IDLista).TipoReserva
            'Label8.Text = oReservas.Item(IDLista).Servicio


        End If
    End Sub
    'Private Sub CargarFicha()
    '    Label1.Text = IDLista
    '    Label6.Text = "Mesa:" & oReservas.Item(IDLista).Mesa
    '    Label7.Text = "Cant:" & oReservas.Item(IDLista).CantidadPax
    '    Label5.Text = oReservas.Item(IDLista).NombreReserva
    '    Label9.Text = oReservas.Item(IDLista).Observaciones
    '    Label4.Text = oReservas.Item(IDLista).TipoReserva
    '    Label8.Text = oReservas.Item(IDLista).Servicio
    'End Sub
    Public Sub CargarFicha(ByVal aIDlista As String)
        'lleno los labels etc
        Label1.Text = aIDlista
        Label6.Text = oReservas.Item(aIDlista).Mesa
        Label7.Text = "Cant:" & oReservas.Item(aIDlista).CantidadPax
        Label5.Text = oReservas.Item(aIDlista).NombreReserva
        Label9.Text = oReservas.Item(aIDlista).Observaciones
        Label4.Text = oReservas.Item(aIDlista).TipoFactura
        Label8.Text = oReservas.Item(aIDlista).Servicio

    End Sub
    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If MsgBox("Salir del sistema?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Me.Dispose()
        End If

    End Sub


    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel1.Click

    End Sub

    Private Sub Acreditar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        IngresarReserva()
    End Sub

    Private Sub IngresarReserva()
        If Not ListView1.SelectedItems.Count = 0 Then ' si no selecciono nada no hace nada
            If (ListView1.SelectedItems(0).SubItems(2).Text) = True Then 'si ya acredito avisa
                If MsgBox("Ya se ingreso x" & ListView1.SelectedItems(0).SubItems(2).Text & ", Reingresar?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Return
                End If
            End If
            oReservas.Item(IDLista).Ingresar()
            CargarComparador()
            CargarListaSinBuscar(ListView1)
            'ListView1.Items(ListView1.FocusedItem.Index).SubItems(2).Text = 1 + ListView1.Items(ListView1.FocusedItem.Index).SubItems(2).Text
            'ListView1.Items(ListView1.FocusedItem.Index).BackColor = Color.LightGreen
            ActulizarIndicadores()
        End If
    End Sub
    Public Sub IngresarReserva(ByVal aidreserva As Integer)
        oReservas.Item(aidreserva).Ingresar()
        CargarComparador()
        CargarListaSinBuscar(ListView1)
        ActulizarIndicadores()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

    End Sub

    Private Sub TableLayoutPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub
    Private Sub RefrescarTodo()
        oReservas.SincronizarBase()
        CargarComparador()
        CargarListaSinBuscar(ListView1)
        ActulizarIndicadores()
    End Sub


    Private Sub Refrescarbt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Refrescarbt.Click
        RefrescarTodo()
        TextBox1.Focus()
        TextBox1.SelectAll()


    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If TextBox1.Text.Length > 0 Then
            TextBox1.Focus()
            '          TextBox1.SelectAll()
        Else
            RefrescarTodo()
        End If


    End Sub
End Class
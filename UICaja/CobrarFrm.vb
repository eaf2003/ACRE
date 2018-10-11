Public Class CobrarFrm
    ' Dim oCobros = Caja.oCobros


    Private Sub CobrarFrm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ComboMoneda_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboMoneda.SelectedIndexChanged
        Label29DescuentoMoneda.Text = "PAR:" & ColMonedaTC(ComboMoneda.Text) & " DESC:" & 100 * ColMonedaDESC(ComboMoneda.Text) & "%"
    End Sub

    Private Sub CargarGrilaPago()
        '        DataGridView1.Rows.Clear()
        DataGridView1.DataSource = Caja.oCobros.VerCobroReserva(Caja.IDLista)
    End Sub

    Private Sub CalcularACobrar(ByVal CantTarjeta As Double, ByVal CantARS As Double, ByVal CantMonExt As Double)

        If IsNumeric(CantTarjeta) And IsNumeric(CantARS) And IsNumeric(CantMonExt) Then
            'devuelve lo que va a cobrar, incluyendo descuentos etc, en ARS
            Dim vTCambio As Double = ColMonedaTC(ComboBox8.Text)
            Dim vDescTotal As Double = CantMonExt - CantMonExt * ColMonedaDESC(ComboBox8.Text)
            '    Dim vMonExtenARS As Double = CantMonExt * vTCambio - vDescTotal
            Dim vMonExtenARS As Double = CantMonExt * vTCambio

            '    TextBox5ACobFinalARS.Text = CantTarjeta + CantARS + vMonExtenARS

        End If

    End Sub

    Private Sub CalcularACobrarMoneda()

        Label34.Text = "PAR:" & ColMonedaTC(ComboBox8.Text) & " DESC:" & 100 * ColMonedaDESC(ComboBox8.Text) & "%"

        Dim vTCambio As Double = ColMonedaTC(ComboBox8.Text)
        Dim vCantMonExt As Double = Label36TotalARS.Text / vTCambio
        Dim vMonExtConDesc As Double = vCantMonExt - vCantMonExt * ColMonedaDESC(ComboBox8.Text)

        TextBox2TotalARS.Text = Math.Round(vMonExtConDesc, 2)

    End Sub
    Public Sub CargarComboMoneda()

        ComboMoneda.DisplayMember = "key"
        ComboMoneda.ValueMember = "value"
        ComboMoneda.DataSource = New BindingSource(ColMonedaTC, Nothing)
        ComboMoneda.SelectedIndex = 0

        ComboBox8.DisplayMember = "key"
        ComboBox8.ValueMember = "value"
        ComboBox8.DataSource = New BindingSource(ColMonedaTC, Nothing)
        ComboBox8.SelectedIndex = 0
    End Sub

    Public Sub cargarComboTarjetas()
        ComboTarjeta.DisplayMember = "key"
        ComboTarjeta.ValueMember = "value"
        ComboTarjeta.DataSource = New BindingSource(ColTarjetaDESC, Nothing)

    End Sub
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        Label4.Text = Caja.IDLista
        Label36TotalARS.Text = Caja.TextBox2TotalARS1.Text

        Me.Text = "Cobrar Reserva : " & Label4.Text

        CargarComboMoneda()
        CalcularACobrarMoneda()
        cargarComboTarjetas()

        TextBox6ACobARS2.Text = 0
        TextBox7ACobMoneda.Text = 0
        TextBox8ACobTarj.Text = 0
        ' Add any initialization after the InitializeComponent() call.

        CargarGrilaPago()
        'oCobros = New BIZL.Cobros(Caja.FechaReservas)
        TextBox6ACobARS2.SelectAll()
        TextBox6ACobARS2.Focus()
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox8.SelectedIndexChanged
        CalcularACobrarMoneda()
    End Sub

    Private Sub ComboTarjeta_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboTarjeta.SelectedIndexChanged
        Label31.Text = " DESC:" & 100 * ColTarjetaDESC(ComboTarjeta.Text) & "%"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Pago()
       
        Me.Dispose()
    End Sub
    Private Sub Pago()
        Dim vSipago = 0
        If TextBox6ACobARS2.Text > 0 Then
            vSipago = 1
            Caja.oCobros.Cobrar(Label4.Text, TextBox6ACobARS2.Text, 1, "EFECTIVO", "ARS", 0)
        End If

        If TextBox7ACobMoneda.Text > 0 Then
            vSipago = 1
            Caja.oCobros.Cobrar(Label4.Text, TextBox7ACobMoneda.Text, ColMonedaTC(ComboMoneda.Text), "EFECTIVO", ComboMoneda.Text, ColMonedaDESC(ComboMoneda.Text))

        End If

        If TextBox8ACobTarj.Text > 0 Then
            vSipago = 1
            Caja.oCobros.Cobrar(Label4.Text, TextBox8ACobTarj.Text, 1, ComboTarjeta.Text, "ARS", ColTarjetaDESC(ComboTarjeta.Text))

        End If
        If vSipago = 1 Then ' si pago algo, le pongo el equiv en pesos
            '            Caja.oReservas.Item(Label4.Text).Pago = CDbl(Label36TotalARS.Text)
            Caja.oReservas.Item(Label4.Text).Pago = vSipago
        End If

        Caja.oReservas.GuardarCambios()
        Caja.CargarComparador()
        Caja.CargarListaSinBuscar(Caja.ListView1)
        Caja.CargarFicha(Label4.Text)

    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()

    End Sub

    Private Sub TextBox6ACobARS2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6ACobARS2.KeyPress
  
        e.Handled = onlyNumbers(e.KeyChar)
    End Sub

    Private Sub TextBox6ACobARS2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6ACobARS2.TextChanged
        If Not IsNumeric(TextBox6ACobARS2.Text) Then
            TextBox6ACobARS2.Text = 0
            TextBox6ACobARS2.SelectAll()
        End If
        If IsNumeric(TextBox7ACobMoneda.Text) And IsNumeric(TextBox6ACobARS2.Text) And IsNumeric(TextBox8ACobTarj.Text) Then
            CalcularFinalenARS(TextBox6ACobARS2.Text, TextBox7ACobMoneda.Text, TextBox8ACobTarj.Text)
        End If
    End Sub

    Private Sub CalcularFinalenARS(ByVal val1ARS As Double, ByVal val2Mon As Double, ByVal val3Tarj As Double)



        Dim vValMonenARS = val2Mon * ColMonedaTC(ComboMoneda.Text) 'paso el valor a ARS

        '  Dim vDescuentoMon = vValMonenARS * ColMonedaDESC(ComboMoneda.Text) 'hago descuento segun la moneda, va por porc
        Dim vdescuentomon = 0
        Dim vDescuentoTar = val3Tarj * ColTarjetaDESC(ComboTarjeta.Text)

        Label5.Text = val1ARS + vValMonenARS - vdescuentomon + val3Tarj - vDescuentoTar


        If Label5.Text < (Label36TotalARS.Text - (Label5.Text + val2Mon * 1.2 * 13.3)) Then
            Label6.Text = "faltan ARS " & (Label36TotalARS.Text - (Label5.Text + val2Mon * 13.3 * 1.2))
        End If
    End Sub

    Public Function onlyNumbers(ByVal KeyChar As Char) As Boolean
        Dim allowedChars As String

        allowedChars = "0123456789."

        If allowedChars.IndexOf(KeyChar) = -1 And (Asc(KeyChar)) <> 8 Then
            Return True
        End If

        Return False
    End Function

    Private Sub TextBox7ACobMoneda_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7ACobMoneda.KeyPress
        e.Handled = onlyNumbers(e.KeyChar)
    End Sub

    Private Sub TextBox7ACobMoneda_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7ACobMoneda.TextChanged
        If Not IsNumeric(TextBox7ACobMoneda.Text) Then
            TextBox7ACobMoneda.Text = 0
            TextBox7ACobMoneda.SelectAll()
        End If

        If IsNumeric(TextBox7ACobMoneda.Text) And IsNumeric(TextBox6ACobARS2.Text) And IsNumeric(TextBox8ACobTarj.Text) Then
            CalcularFinalenARS(TextBox6ACobARS2.Text, TextBox7ACobMoneda.Text, TextBox8ACobTarj.Text)
        End If

    End Sub

    Private Sub TextBox8ACobTarj_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox8ACobTarj.KeyPress
        e.Handled = onlyNumbers(e.KeyChar)
    End Sub

    Private Sub TextBox8ACobTarj_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8ACobTarj.TextChanged
        If Not IsNumeric(TextBox8ACobTarj.Text) Then
            TextBox8ACobTarj.Text = 0
            TextBox8ACobTarj.SelectAll()
        End If

        If IsNumeric(TextBox7ACobMoneda.Text) And IsNumeric(TextBox6ACobARS2.Text) And IsNumeric(TextBox8ACobTarj.Text) Then
            CalcularFinalenARS(TextBox6ACobARS2.Text, TextBox7ACobMoneda.Text, TextBox8ACobTarj.Text)
        End If

    End Sub

    Private Sub DataGridView1_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DataGridView1.RowsRemoved

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Caja.oCobros.UpdateadDatosBindeados()
        DataGridView1.EndEdit()
    End Sub
End Class
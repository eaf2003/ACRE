Public Class Walkin

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose()
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim winID As Integer
        winID = Caja.oReservas.AgregarWalkIn(TextBox1.Text, ComboBox1.Text, "WI- Walk In", MaskedTextBox2.Text, MaskedTextBox1.Text, 0, ComboBox2.Text, TextBox4.Text)
        Caja.CargarComparador()
        Caja.CargarListaSinBuscar(Caja.ListView1)
        Caja.CargarFicha(winID)
        Caja.ImprimirToken()
        MsgBox("Acredito: " & winID)

        Me.Dispose()

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'LLENO COMBOS, LOS DEJO FIJOS SO FAR


        ComboBox2.DisplayMember = "key"
        ComboBox2.ValueMember = "value"
        ComboBox2.DataSource = New BindingSource(ColTipoFactura, Nothing)
        ComboBox2.AutoCompleteSource = AutoCompleteSource.ListItems
        ComboBox2.SelectedIndex = 1


        ComboBox1.DataSource = Caja.ComboBox1Servicios.Items
        ComboBox1.ValueMember = "servicio"
        ComboBox1.Text = "V- Cena Show VIP"

        MaskedTextBox1.Text = 2
        MaskedTextBox2.Text = 0


        'ComboBox1.Items.Add("PV- Cena Show Palco Vip")
        'ComboBox1.Items.Add("PR- Cena Show Premium")
        'ComboBox1.Items.Add("VV- Cena Show VIP")
        'ComboBox1.Items.Add("EJ- Cena Show Ejecutiva")
        'ComboBox1.Items.Add("SV- Solo Show VIP")
        'ComboBox1.Items.Add("SO- Solo Show Ejecutiva")
        'ComboBox1.Items.Add("SP- Solo Palco Vip")
        'ComboBox1.Items.Add("SC- Show Sin Consumicion")
        'ComboBox1.Items.Add("TE- Terraza PB Fiestas")
        'ComboBox1.Items.Add("LA- Last Minute")





    End Sub

    Private Sub MaskedTextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MaskedTextBox1.GotFocus
        Me.MaskedTextBox1.SelectAll()
    End Sub

    Private Sub MaskedTextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MaskedTextBox1.KeyPress
        If (e.KeyChar = " ") Then
            e.KeyChar = ""
        End If

        e.Handled = onlyNumbers(e.KeyChar)
    End Sub

    Private Sub MaskedTextBox1_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs) Handles MaskedTextBox1.MaskInputRejected

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Button2.Enabled = True
    End Sub

    Public Function onlyNumbers(ByVal KeyChar As Char) As Boolean
        Dim allowedChars As String

        allowedChars = "0123456789"

        If allowedChars.IndexOf(KeyChar) = -1 And (Asc(KeyChar)) <> 8 Then
            Return True
        End If

        Return False
    End Function

    Private Sub MaskedTextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MaskedTextBox2.GotFocus
        Me.MaskedTextBox2.SelectAll()
    End Sub

    Private Sub MaskedTextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MaskedTextBox2.KeyPress
        If (e.KeyChar = " ") Then
            e.KeyChar = ""
        End If

        e.Handled = onlyNumbers(e.KeyChar)
    End Sub
End Class
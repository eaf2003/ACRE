Public Class CambioServicio

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.


        ComboBox1.Items.Add("UPGRADE")
        ComboBox1.Items.Add("ERROR DE CARGA")
        ComboBox1.Items.Add("OTRO")
        ComboBox1.SelectedIndex = 0

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            ComboBox2.Visible = True
            Label2.Visible = True
        Else
            ComboBox2.Visible = False
            Label2.Visible = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Caja.oReservas.Item(Caja.Label1.Text).Servicio = Caja.ComboBox1Servicios.Text
        Caja.Label8.Text = Caja.ComboBox1Servicios.Text
        Caja.Label7.Text = Caja.MaskedTextBox1Cant.Text
        Me.Hide()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub
End Class
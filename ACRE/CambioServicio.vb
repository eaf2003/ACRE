Imports System.Threading
Imports System.Globalization
Public Class CambioServicio

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'por cosas de fechas y demas seteo la cultura fija
        '  Thread.CurrentThread.CurrentCulture = New CultureInfo("es-AR")
        ' Thread.CurrentThread.CurrentUICulture = New CultureInfo("es-AR")

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
        Acreditaciones.oReservas.Item(Acreditaciones.Label1.Text).Servicio = Acreditaciones.ComboBox1.Text
        Acreditaciones.Label8.Text = Acreditaciones.ComboBox1.Text
        Acreditaciones.Label7.Text = Acreditaciones.MaskedTextBox1.Text
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
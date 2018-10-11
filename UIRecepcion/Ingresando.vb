Public Class Ingresando

    Public Sub New()

        ' This call is required by the Windows Form Designer.

        InitializeComponent()

        Me.Text = Me.Text & " " & Recepcion.Label1.Text
        ProgressBar1.Maximum = 100
        ' Add any initialization after the InitializeComponent() call.
        Timer1.Interval = 100 'velocidad , frecuencia del timer
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        ProgressBar1.Value = ProgressBar1.Value + 10

        If ProgressBar1.Value = ProgressBar1.Maximum Then

            Me.Dispose()
        End If

    End Sub
End Class
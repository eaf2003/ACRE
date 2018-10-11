
Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        
        frmBuscaCliente.Show()

    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

     

    End Sub

    Private Sub Form1_MaximizedBoundsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MaximizedBoundsChanged

    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        Dim cfg As New Datos.Configuracion

        TextBox1.Text = (cfg.DevolverValor("ExcelPath"))

        ' Add any initialization after the InitializeComponent() call.



    End Sub
End Class

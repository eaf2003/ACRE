Public Class frmBuscaCliente

    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub
    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        Dim oXLSFILE As New Datos.BaseXLS

        oXLSFILE.Conectar(Form1.TextBox1.Text)

        DataGridView1.DataSource = oXLSFILE.TomarDatos("select * from [Hoja1$]")

        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class
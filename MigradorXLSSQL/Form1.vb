Imports System.Threading
Imports System.Globalization
Public Class Form1
    Public Shared cfg As New ACRE.Datos.Configuracion



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        
        Migrador.Show()

    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

     

    End Sub

    Private Sub Form1_MaximizedBoundsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MaximizedBoundsChanged

    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        InitCultura()

        TextBox1.Text = cfg.DevolverValor("ExcelPath", "migrador.xml")
        TextBox2.Text = cfg.DevolverValor("ExcelHoja", "migrador.xml")
        TextBox3.Text = cfg.DevolverValor("ExcelCampoID", "migrador.xml")

        ' Add any initialization after the InitializeComponent() call.



    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim f As New OpenFileDialog
        f.Filter = "Excel files | *.xls"
        f.InitialDirectory = Application.ExecutablePath

        If (f.ShowDialog =Windows.Forms.DialogResult.OK) Then
            If (f.FileName <> "" And f.CheckFileExists = True) Then
                TextBox1.Text = f.FileName
            End If
        End If
				
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CenaPasada.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TipoDeCambioMod.Show()
    End Sub
End Class

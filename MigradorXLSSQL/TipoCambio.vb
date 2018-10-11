Imports System.Threading
Imports System.Globalization
Public Class TipoDeCambioMod
    'Dim oBaseSql As New Datos.BaseSQL
    'Dim dtsql As New DataTable

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()


 
        ' Add any initialization after the InitializeComponent() call.
        Try

            Dim oSetup As New BIZL.SetUp

            '        dtsql = oBaseSql.TomarDatos(qry)
            DataGridView1.DataSource = oSetup.oDTMoneda

            'SET FECHA

            'If MsgBox("DEFINIR FECHA DE LA CENA DE HOY COMO:" & Form1.DateTimePicker1.Text, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            '    Dim sqlCmd2 As New SqlClient.SqlCommand
            '    sqlCmd2.Parameters.Add(New SqlClient.SqlParameter("@Fecha", SqlDbType.SmallDateTime))
            '    Dim strSQLUPDCmd2 As String = "UPDATE RESERVASSETUP SET FECHA=@Fecha WHERE PARAMETRO='FECHACENA' "
            '    sqlCmd2.CommandText = strSQLUPDCmd2

            '    sqlCmd2.Parameters("@Fecha").Value = Form1.DateTimePicker1.Value.Date
            '    oBaseSql.ExecutarQuery(sqlCmd2)

            'Else
            '    Return

            'End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Migrador_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Splitter1_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs)

    End Sub

    Private Sub OpenToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub HelpToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Extras.Click





    End Sub

    Private Sub Salir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Salir.Click
        Form1.Dispose()
        Me.Dispose()

    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Guardar.Click
        'MsgBox("hola", MsgBoxStyle.Exclamation)
        '     oBaseSql.UpdateDatosBindeados()
        DataGridView1.EndEdit() 'sin esto no updatea el ultimo valor editado
    End Sub
End Class

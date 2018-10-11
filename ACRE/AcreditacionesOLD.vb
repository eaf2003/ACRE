
Public Class Form2

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub HideResultado()
        ListBox1.Visible = False
    End Sub

    Private Sub AutoBusqueda(ByVal ListViewRes As ListView, ByVal TextBoxBusq As TextBox)
        ' AUTO BUSQUEDA BUSCA EL STRING EN EL TEXTO
        ListViewRes.Items.Clear()
        Dim str1(1) As String

        If (TextBoxBusq.Text.Length > 1) Then 'si escribe mas de 1 caracteres empieza a buscar

            If (TextBoxBusq.Text.Length = 0) Then
                'HideResultado() 'oculto el list
                Return
            End If
            'For Each s In TextBox1.AutoCompleteCustomSource


            For Each s As KeyValuePair(Of Integer, String) In AComP2
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
                'otros controles a llenar cuando encontro de una
                Label3.Text = str1(0)

            End If
            '''''

        End If
        '      Return AComP2.Count

    End Sub

    Private Sub AutoBusquedaOLD()
        ' AUTO BUSQUEDA BUSCA EL STRING EN EL TEXTO
        ListBox1.Items.Clear()
        ListView1.Items.Clear()
        Dim str1(1) As String

        If (TextBox1.Text.Length > 1) Then 'si escribe mas de 1 caracteres empieza a buscar

            If (TextBox1.Text.Length = 0) Then
                HideResultado() 'oculto el list
                Return
            End If
            'For Each s In TextBox1.AutoCompleteCustomSource


            For Each s As KeyValuePair(Of Integer, String) In AComP2
                If s.Value().Contains(TextBox1.Text) Or s.Key.ToString.Contains(TextBox1.Text) Then
                    '  Console.WriteLine("hallo:" + s.Key.ToString + s.Value)

                    ListBox1.Items.Add(s.Key.ToString + s.Value)
                    ListBox1.Visible = True

                    'para llenar el listview


                    str1(0) = s.Key.ToString
                    str1(1) = s.Value

                    Dim LVi1 As New ListViewItem(str1)
                    ListView1.Items.Add(LVi1)

                End If
            Next

            'si solo encontro uno automaticamente llego el textbox con ese valor
            If ListView1.Items.Count = 1 Then
                TextBox1.Text = str1(1)
                Label3.Text = str1(0)
                TextBox1.SelectAll()
            End If
            '''''

        End If
        '      Return AComP2.Count
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress

    End Sub

    Private Sub TextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyUp

        AutoBusqueda(ListView1, TextBox1)
   
    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TableLayoutPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'EAF LLENAR AUTOCOMPLETE 
        Dim BD1 As New Datos.BaseXLS
        BD1.Conectar(Form1.TextBox1.Text)
        Dim TB1 As New DataTable
        'TB1 = BD1.TomarDatos("select * from [Hoja1$] where id is not null")
        TB1 = BD1.TomarDatos("select * from [U-Res-Consulta X EStado Diario $] where mesa is not null")
        For Each rw As DataRow In TB1.Rows
            'AComP.Add(rw.Item("RazonSocial"))

            '  Console.WriteLine(rw.Item("RazonSocial"))
            AComP2.Add(rw.Item("Nº Rva"), UCase(rw.Item("Nombre de Reserva"))) 'meto key nro(debe ser unico sino aexplota) y valor string
        Next


        TB1.Dispose()



    End Sub

    Private Sub ListBox1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.Enter


    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedIndex > -1 Then
            TextBox1.Text = ListBox1.Items(ListBox1.SelectedIndex).ToString()
            'HideResultado()
        End If
    End Sub


    Private Sub ListView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        If ListView1.SelectedItems.Count > 0 Then

            TextBox1.Text = ListView1.SelectedItems(0).SubItems(1).Text
            Label3.Text = ListView1.SelectedItems(0).SubItems(0).Text.ToString
            'HideResultado()

        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim item1 As ListViewItem = ListView1.FindItemWithText("ho", True, 0, False)
        If (item1 IsNot Nothing) Then
            MessageBox.Show("Calling FindItemWithText passing 'brack': " _
                & item1.ToString())
        Else
            MessageBox.Show("NOOCalling FindItemWithText passing 'brack': null")
        End If

    End Sub
End Class
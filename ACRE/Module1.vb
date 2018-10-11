Imports System.Threading
Imports System.Globalization
Module Module1
    'EAF START collecion de elementos a buscar
    Public AComP As AutoCompleteStringCollection
    Public AComP2 As New Dictionary(Of Integer, String)
    Public AcomP3 As New Dictionary(Of Int32, ColAComp3)
    Public ColTipoRes As New Dictionary(Of String, String)
    Public ColTipoFactura As New Dictionary(Of String, String)

    Public ColServicioDescripcion As New Dictionary(Of String, String)

    Dim oCfg As New Datos.Configuracion

    Public Sub InitCultura()
        Dim STRlocale As String = oCfg.DevolverValor("locale", "base.xml")
        Thread.CurrentThread.CurrentCulture = New CultureInfo(STRlocale)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(STRlocale)
        Dim Cu As New CultureInfo(STRlocale)
        '    Cu.DateTimeFormat.ShortDatePattern = "YYYYMMDD" '"dd-MM-yyyy"
    End Sub
    Public Sub ShowCurrentCulture()
        MsgBox("Culture of {0} in application domain {1}: {2}" & Thread.CurrentThread.Name & AppDomain.CurrentDomain.FriendlyName & CultureInfo.CurrentCulture.Name & CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern)

    End Sub

    ' Put the following code before InitializeComponent()
    ' Sets the culture to French (France)

End Module

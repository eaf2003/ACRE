Imports System.Threading
Imports System.Globalization
Module Module1
    'EAF START collecion de elementos a buscar
    Public AComP As AutoCompleteStringCollection
    Public AComP2 As New Dictionary(Of Integer, String)
    Public AcomP3 As New Dictionary(Of Int32, ColAComp3)
    Dim oCfg As New Datos.Configuracion

    Public Sub InitCultura()
        Dim STRlocale As String = oCfg.DevolverValor("locale", "base.xml")
        Thread.CurrentThread.CurrentCulture = New CultureInfo(STRlocale)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(STRlocale)
        Dim Cu As New CultureInfo(STRlocale)
        '    Cu.DateTimeFormat.ShortDatePattern = "YYYYMMDD" '"dd-MM-yyyy"
    End Sub


End Module

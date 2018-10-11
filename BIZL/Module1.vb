Imports System.Threading
Imports System.Globalization
Module Module1
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

End Module

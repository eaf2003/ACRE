Imports System.Xml

Public Class Configuracion


    Function DevolverValor(ByVal NombreKey As String) As String

        Try


            Dim xtr As New XmlTextReader("config.xml")
            Do While xtr.Read
                If xtr.NodeType = XmlNodeType.Element Then
                    ' Publisher names are inserted as text immediately
                    ' after an element named pub_name.
                    If xtr.Name = NombreKey Then
                        ' Move to the next element, and display its value.
                        xtr.Read()
                        ' Console.WriteLine(xtr.Value)
                        Return xtr.Value

                    End If
                End If
            Loop
            xtr.Close()
            xtr = Nothing




        Catch ex As Exception
            MsgBox("Base.Configuracion   " & ex.Message)
        End Try
        Return -1
    End Function

    Function DevolverValor(ByVal NombreKey As String, ByVal XMLfile As String) As String
        Dim xtr As New XmlTextReader(XMLfile)
        Do While xtr.Read
            If xtr.NodeType = XmlNodeType.Element Then
                ' Publisher names are inserted as text immediately
                ' after an element named pub_name.
                If xtr.Name = NombreKey Then
                    ' Move to the next element, and display its value.
                    xtr.Read()
                    ' Console.WriteLine(xtr.Value)
                    Return xtr.Value

                End If
            End If
        Loop
        xtr.Close()
        xtr = Nothing
        Return -1

    End Function

End Class


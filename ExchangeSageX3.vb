Imports System.Xml
Imports System.IO

Module Module1

    Sub Main()

        Try

            Dim fecha As String
            Dim gbp As String
            Dim usd As String
            Dim chf As String
            Dim nuevaFech As String
            Dim xmlReader As XmlReader
            Dim valores(0) As String
            Dim i As Integer

            xmlReader = New Xml.XmlTextReader("http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml")

            While xmlReader.Read
                Select Case xmlReader.NodeType
                    Case Xml.XmlNodeType.Element
                        If xmlReader.AttributeCount > 0 Then
                            While xmlReader.MoveToNextAttribute
                                valores(i) = xmlReader.Value
                                ReDim Preserve valores(i + 1)
                                Console.WriteLine(valores(i))
                                i = i + 1
                            End While
                        End If
                End Select
            End While

            xmlReader.Close()

            fecha = valores(2)
            For i = 0 To valores.Length - 1
                Select Case valores(i)
                    Case "USD"
                        usd = valores(i + 1)
                    Case "GBP"
                        gbp = valores(i + 1)
                    Case "CHF"
                        chf = valores(i + 1)
                    Case Else
                End Select
            Next

            nuevaFech = fecha.Substring(8, 2)
            nuevaFech += fecha.Substring(5, 2)
            nuevaFech += fecha.Substring(2, 2)

            Dim swEscritor As StreamWriter
            swEscritor = New StreamWriter("C:\SAGE\AUTOIMPORT\change.txt", False)

            swEscritor.WriteLine("""EUR""" & vbTab & """GBP""" & vbTab & nuevaFech & vbTab & "1" & vbTab & gbp)
            swEscritor.WriteLine("""EUR""" & vbTab & """USD""" & vbTab & nuevaFech & vbTab & "1" & vbTab & usd)
            swEscritor.WriteLine("""EUR""" & vbTab & """CHF""" & vbTab & nuevaFech & vbTab & "1" & vbTab & chf)

            swEscritor.Close()

            Console.WriteLine("Archivo de divisas creado con exito")
            
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

    End Sub

End Module

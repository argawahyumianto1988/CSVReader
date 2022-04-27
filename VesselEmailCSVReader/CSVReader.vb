Imports Microsoft.VisualBasic.FileIO
Imports System.IO

Public Class CSVReader

    'mendapatkan data CSV dalam bentuk array of string,index array urut berdasarkan kolom dari kiri ke kanan
    Public Shared Function GetCSVData(ByVal path As String) As List(Of String())

        Dim list As New List(Of String())

        'path = "../../WatchReport_20131210_0600.csv"
        Try
            Using parser As New TextFieldParser(path)
                'parser.CommentTokens = New String() {"#"}
                'parser.SetDelimiters(New String() {";"})
                parser.HasFieldsEnclosedInQuotes = True
                parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
                parser.Delimiters = New String() {","}

                'don't Skip over header line.
                'parser.ReadLine()

                While Not parser.EndOfData
                    Dim fields As String() = parser.ReadFields()
                    list.Add(fields)
                End While

            End Using
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        Return list

    End Function

    'mendapatkan data CSV dalam bentuk dictionary,key berdasarkan string header dari kolom
    'untuk header yang tidak lengkap maka otomatis akan dibuatkan berdasarkan urut index (misal string header jumlah 4 tapi datanya ada 10 kolom)
    Public Shared Function GetCSVDatabyDictionary(ByVal path As String) As List(Of Dictionary(Of String, String))

        Dim list As New List(Of Dictionary(Of String, String))
        Dim listTemp As Dictionary(Of String, String)

        'path = "../../WatchReport_20131210_0600.csv"
        Try
            Using parser As New TextFieldParser(path)
                'parser.CommentTokens = New String() {"#"}
                'parser.SetDelimiters(New String() {";"})
                parser.HasFieldsEnclosedInQuotes = True
                parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
                parser.Delimiters = New String() {","}

                ' Skip over header line.
                'parser.ReadLine()
                Dim stringHeader As String() = parser.ReadFields()

                While Not parser.EndOfData
                    Dim fields As String() = parser.ReadFields()
                    'oke pake ini

                    listTemp = New Dictionary(Of String, String)
                    Dim iter As Integer = 0
                    For Each currentField As String In fields
                        Console.WriteLine(currentField)
                        Try
                            listTemp.Add(stringHeader(iter), currentField)
                        Catch ex As Exception
                            listTemp.Add(iter, currentField)
                            'Console.WriteLine(ex.Message)
                        End Try
                        iter += 1
                    Next

                    list.Add(listTemp)

                End While

            End Using
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        Return list

    End Function

End Class

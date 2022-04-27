Imports System.Text

Class CSVWriter

    Sub CreateCSVHeader(ByVal filename As String, ByVal headerString As List(Of String))
        Try
            Dim objWriter As IO.StreamWriter = IO.File.AppendText(filename)

            If IO.File.Exists(filename) Then
                objWriter.WriteLine(CsvHeader(headerString))
            Else
                MsgBox(" File Does not exist")
            End If

            objWriter.Close()
            objWriter.Dispose()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

    End Sub

    Sub CreateCSVAppend(ByVal filename As String, ByVal fillString As List(Of String))
        Try
            Dim objWriter As IO.StreamWriter = IO.File.AppendText(filename)

            If IO.File.Exists(filename) Then
                objWriter.WriteLine(CsvContent(fillString))
            Else
                MsgBox(" File Does not exist")
            End If

            objWriter.Close()
            objWriter.Dispose()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

    End Sub

    Function CsvHeader(ByVal headerString As List(Of String)) As String
        Dim CsvLine As New StringBuilder
        Dim headerLength = headerString.Count

        Try
            For Each iter As String In headerString
                If headerLength > 1 Then
                    CsvLine.Append(iter + ",")
                Else
                    CsvLine.Append(iter)
                End If

                headerLength -= 1
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        Return CsvLine.ToString
    End Function

    Function CsvContent(ByVal fillString As List(Of String)) As String
        Dim CsvLine As New StringBuilder
        Dim contentLength = fillString.Count

        Try
            For Each iter As String In fillString
                If contentLength > 1 Then
                    CsvLine.Append(iter + ",")
                Else
                    CsvLine.Append(iter)
                End If

                contentLength -= 1
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        Return CsvLine.ToString
    End Function

    'untuk mengatasi data yang ada komanya
    Function EncodeComma(ByVal value As String) As String
        Return """" & value & """"
    End Function

    Sub CreateDirectory(ByVal dir As String)
        If Not System.IO.Directory.Exists(dir) Then
            System.IO.Directory.CreateDirectory(dir)
        End If
    End Sub

End Class


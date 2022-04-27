Imports System.IO

Public Class LogHelper
    Public Shared Sub InsertLog(ByVal message As String)
        Dim dateNow = DateTime.Now.Day & "-" & DateTime.Now.Month & "-" & DateTime.Now.Year
        If Not System.IO.Directory.Exists("Log\") Then
            System.IO.Directory.CreateDirectory("Log\")
        End If

        'check the file
        'Dim fs As FileStream = New FileStream("\Log\LogMessage.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        'Dim s As StreamWriter = New StreamWriter(fs)
        's.Close()
        'fs.Close()

        'log it
        Dim fs1 As FileStream = New FileStream("Log\LogMessage" & dateNow & ".txt", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)
        s1.WriteLine(message)
        s1.Close()
        fs1.Close()
    End Sub
End Class

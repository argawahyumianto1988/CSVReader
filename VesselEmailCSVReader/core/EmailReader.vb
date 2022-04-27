Imports System.IO

Public Class EmailReader
    Public WithEvents POP3Session As OSPOP3_Plus.Session    
    Private host As String
    Private port As Integer
    Private username As String
    Private password As String    

    Public Sub New()

    End Sub

    Public Sub Configure(ByVal host As String, ByVal port As Integer, ByVal usern As String, ByVal passw As String)
        Me.host = host
        Me.port = port
        Me.username = usern
        Me.password = passw
    End Sub

    Public Function GetMailCount() As Integer
        Try
            Return POP3Session.GetMessageCount()
        Catch ex As Exception
            Dim a As String = ex.Message
            Return 0
        End Try
    End Function

    Public Function GetMailSize() As Integer
        Try
            Return POP3Session.GetMailboxSize()
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function GetMessageList() As Dictionary(Of Integer, OSPOP3_Plus.Message)
        Dim message As OSPOP3_Plus.Message
        Dim MessageList As New Dictionary(Of Integer, OSPOP3_Plus.Message)
        Dim dateNow = DateTime.Now.Day & "-" & DateTime.Now.Month & "-" & DateTime.Now.Year
        Dim dir As String = "Email\Message\" & dateNow & "\"
        'diambil dari email sender 
        Dim vesselName As String = ""

        For Each oMLE As OSPOP3_Plus.MessageListEntry In POP3Session.GetMessageList()
            'pada server dihapus
            'True jika email ingin dihapus setelah diambil
            message = POP3Session.GetMessage(oMLE.ID, True)
            'message = POP3Session.GetMessage(oMLE.ID)

            vesselName = message.Sender.Address.Split("@")(0)

            Dim emailSavedDirectory = DownloadEmailMetadata(dir + vesselName + "\", message.UIDL)
            message.Save(emailSavedDirectory)
            'Console.WriteLine(message.HTMLBody)

            MessageList.Add(oMLE.ID, message)
        Next
        Return MessageList

    End Function

    Private Function GetPOP3Status()
        Return POP3Session.Status
    End Function

    Public Function OpenConnectionMailServer() As Boolean

        Try
            'configure POP3 Session
            POP3Session = New OSPOP3_Plus.Session
            POP3Session.UseSSL = True
            POP3Session.Login = username
            POP3Session.PortNumber = 995
            POP3Session.ServerName = host
            POP3Session.Password = password

            'konek POP3 server            
            POP3Session.OpenPOP3()
            'POP3Session.OpenPOP3(Me.host, Me.port, Me.username, Me.password)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        Return False

    End Function

    Public Sub CloseConnectionMailServer()
        Try
            POP3Session.ClosePOP3()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub oSession_Connected() Handles POP3Session.Connected

    End Sub

    Private Sub oSession_Closed() Handles POP3Session.Closed

    End Sub

    Public Function DownloadFileAttachment(ByVal dir As String, ByVal filename As String) As String

        'Dim dateNow = DateTime.Now.Day & "-" & DateTime.Now.Month & "-" & DateTime.Now.Year
        'Dim dir As String = "Attachment\" & dateNow & "\"

        If Not System.IO.Directory.Exists(dir) Then
            System.IO.Directory.CreateDirectory(dir)
        End If

        Try
            Dim fileToSave As String = dir & filename
            Dim fs1 As FileStream = New FileStream(fileToSave, FileMode.OpenOrCreate, FileAccess.ReadWrite)
            fs1.Close()
        Catch ex As Exception
            Console.WriteLine("File sudah ada : " & ex.Message)
        End Try

        Return dir & filename
    End Function

    Public Function DownloadEmailMetadata(ByVal dir As String, ByVal filename As String) As String

        'Dim dateNow = DateTime.Now.Day & "-" & DateTime.Now.Month & "-" & DateTime.Now.Year
        'Dim dir As String = "Email\" & dateNow & "\"

        If Not System.IO.Directory.Exists(dir) Then
            System.IO.Directory.CreateDirectory(dir)
        End If

        Try
            Dim fileToSave As String = dir & filename
            Dim fs1 As FileStream = New FileStream(fileToSave, FileMode.OpenOrCreate, FileAccess.ReadWrite)
            fs1.Close()
        Catch ex As Exception
            Console.WriteLine("File sudah ada : " & ex.Message)
        End Try

        Return dir & filename
    End Function

    Private Sub oSession_StatusChanged(ByVal Status As String, ByVal StatusType As OSPOP3_Plus.Session.StatusTypeConstants) Handles POP3Session.StatusChanged
        If StatusType = OSPOP3_Plus.Session.StatusTypeConstants.stPOP3Request Then
            Console.WriteLine("Status : " & Status)
        End If
        '        Case OSPOP3_Plus.Session.StatusTypeConstants.stPOP3Request : sPrompt = "< "
        '        Case OSPOP3_Plus.Session.StatusTypeConstants.stPOP3Response : sPrompt = "> "
        '        Case OSPOP3_Plus.Session.StatusTypeConstants.stError : sPrompt = "! "
        '        Case OSPOP3_Plus.Session.StatusTypeConstants.stState : sPrompt = "# "
        '        Case Else : sPrompt = "? "
        '    End Select

    End Sub

End Class


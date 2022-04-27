Imports System.IO
Imports System.Threading
Imports System.Data.OleDb

Module Main

    Sub Main()

        'Console.WriteLine(GetDecimalFromNMEA(712.31809059016, 1.0))
        RemoveLogFile(15)
        'Console.WriteLine(GetDecimalFromNMEA(11243.6697674616, 0.0))
        'konfigurasi POP3 MailServer
        Dim er As New EmailReader()
        ''er.Configure("172.16.0.115", 995, "user.training01", "mrt123")
        'er.Configure("172.16.0.215", 995, "rms.meratus", "mrt123")
        er.Configure("outlook.office365.com", 995, "admin.rms@meratusline.com", "rms@dmin")

        ' While True
        'Dim tes As String = """GPS-file_230614-0035-.csv"""
        'Console.WriteLine(tes.Substring(1, tes.Length - 2))
        'Console.ReadLine()
      
        er.OpenConnectionMailServer()
        'Console.WriteLine()
        'Console.WriteLine()
        'Console.WriteLine()
        'Console.WriteLine()
        'Console.WriteLine(er.GetMailCount())
        'Console.WriteLine(er.GetMailSize())
        'Console.WriteLine()
        'Console.WriteLine()

        'Mendapatkan data semua email masuk(inbox)
        Dim listMessage As Dictionary(Of Integer, OSPOP3_Plus.Message) = er.GetMessageList()
        Dim fileToSave As String = ""
        Dim pair As KeyValuePair(Of Integer, OSPOP3_Plus.Message)
        Dim vesselName As String = ""
        Dim dateNow = DateTime.Now.Day & "-" & DateTime.Now.Month & "-" & DateTime.Now.Year
        Dim dirDefault As String = "C:\GPS\Projek_Server_Febri\VesselEmailCSVReader\VesselEmailCSVReader\bin\Debug\Email\Attachment\"
        Dim dirDefaultRMS As String = "C:\GPS\Projek_Server_Febri\VesselEmailCSVReader\VesselEmailCSVReader\bin\Debug\Email\Attachment\"
        Dim dirToVesselArchive As String = ""
        Dim dirToUploadDatabase As String = ""

        For Each pair In listMessage
            'jika ada attachment dan subjeknya GPS
            'meratussemarang di-skip dulu
            If pair.Value.Attachments.Count > 0 And (pair.Value.Subject.Contains("GPS") Or pair.Value.Subject.Contains("RMS")) Then
                'If pair.Value.Attachments.Count > 0 And (pair.Value.Subject.Contains("GPS") Or pair.Value.Subject.Contains("RMS")) And pair.Value.Sender.Address.Split("@")(0) <> "meratus.sumba" And pair.Value.Sender.Address.Split("@")(0) <> "meratus.batam" And pair.Value.Sender.Address.Split("@")(0) <> "meratus.kampar" Then 'pair.Value.ContentTransferEncoding = "base64" Then 'pair.Value.Sender.Address.Split("@")(0) <> "MeratusSemarang" Then
                Console.WriteLine("Subject :" & pair.Value.Subject)
                Console.WriteLine("Sender  :" & pair.Value.Sender.Address)
                Console.WriteLine("Date    :" & pair.Value.DateSent)

                Dim temp As String = ""
                temp = pair.Value.ContentTransferEncoding
                vesselName = pair.Value.Sender.Address.Split("@")(0)

                'For Each header As OSPOP3_Plus.Header In pair.Value.Headers
                '    temp &= header.Name & " "
                'Next
                'Console.WriteLine("Header :" & temp)

                temp = ""
                For Each att As OSPOP3_Plus.Attachment In pair.Value.Attachments
                    If att.ContentTransferEncoding = "base64" Then  'hanya base64
                        If att.Filename <> "" Then
                            temp &= "-" & att.Filename & " "
                            Dim aaa As String = att.Filename.Substring(att.Filename.Length - 4, 3)
                            Dim fsf As String = att.ContentType & att.ContentDisposition & att.Body & att.AttachmentName
                            Dim xxx As String = att.Filename.Substring(1, 8)
                            'jika csv (ato mungkin file lain juga :D)
                            If att.Filename.Substring(att.Filename.Length - 4, 3) = "csv" Then
                                'If att.ContentType = "application/octet-stream" Or att.ContentType = "application/vnd.ms-excel" Or att.ContentType = "text/csv" Then
                                'add by yna 07.12.2015 for RMS
                                If att.Filename.Substring(1, 8) = "KW_METER" Or att.Filename.Substring(1, 8) = "kw_meter" Then
                                    dirToVesselArchive = dirDefaultRMS & dateNow & "\" & vesselName & "\"
                                    'perkecualian karimata
                                    If pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.karimata" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.waigeo" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.manado" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.medan3" Then
                                        dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS\"
                                    Else
                                        dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS_mka\"
                                    End If
                                    Dim a As String = att.Filename.Substring(1, att.Filename.Length - 2)
                                    'disimpan 2 kali satunya untuk arsip satunya untuk diproses ke database kemudian dihapus (tempd_download folder)
                                    fileToSave = er.DownloadFileAttachment(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)

                                    fileToSave = er.DownloadFileAttachment(dirToUploadDatabase, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)

                                    If pair.Value.Sender.Address.Split("@")(0) = "rms.meratus.karimata" Or pair.Value.Sender.Address.Split("@")(0) = "rms.meratus.waigeo" Then
                                        fileToSave = er.DownloadFileAttachment("C:\yusuf\nlab\", pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                        att.Save(fileToSave)
                                    End If
                                    'end of add by yna 07.12.2015 for RMS
                                    'add for encrypt data by yna 17.09.2018
                                ElseIf att.Filename.Substring(1, 9) = "eKW_METER" Or att.Filename.Substring(1, 9) = "ekw_meter" Then
                                    dirToVesselArchive = dirDefaultRMS & dateNow & "\" & vesselName & "\"
                                    'perkecualian benoa
                                    If pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.karimata" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.waigeo" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.manado" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.medan3" Then
                                        dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS\"
                                    Else
                                        dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS_mka\"
                                    End If
                                    dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS_encrypt\"
                                    Dim a As String = att.Filename.Substring(1, att.Filename.Length - 2)
                                    'disimpan 2 kali satunya untuk arsip satunya untuk diproses ke database kemudian dihapus (tempd_download folder)
                                    fileToSave = er.DownloadFileAttachment(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)

                                    fileToSave = er.DownloadFileAttachment(dirToUploadDatabase, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)

                                    If pair.Value.Sender.Address.Split("@")(0) = "rms.meratus.karimata" Then
                                        fileToSave = er.DownloadFileAttachment("C:\yusuf\nlab\", pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                        att.Save(fileToSave)
                                    End If
                                    'end of add for encrypt data by yna 17.09.2018
                                ElseIf att.Filename.Substring(1, 3) = "RMS" Or att.Filename.Substring(1, 3) = "rms" Then
                                    dirToVesselArchive = dirDefaultRMS & dateNow & "\" & vesselName & "\"
                                    dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS_new\"
                                    Dim a As String = att.Filename.Substring(1, att.Filename.Length - 2)
                                    'disimpan 2 kali satunya untuk arsip satunya untuk diproses ke database kemudian dihapus (tempd_download folder)
                                    fileToSave = er.DownloadFileAttachment(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)

                                    fileToSave = er.DownloadFileAttachment(dirToUploadDatabase, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)
                                    'end of add for encrypt data by yna 17.09.2018
                                Else
                                    'add by yna 25.07.2018 delete prefix rms.
                                    vesselName = vesselName.Replace("rms.", "")
                                    'end of add by yna 25.07.2018 delete prefix rms.

                                    dirToVesselArchive = dirDefault & dateNow & "\" & vesselName & "\"
                                    dirToUploadDatabase = dirDefault & dateNow & "\temp_download\"
                                    Dim a As String = att.Filename.Substring(1, att.Filename.Length - 2)
                                    'disimpan 2 kali satunya untuk arsip satunya untuk diproses ke database kemudian dihapus (tempd_download folder)
                                    fileToSave = er.DownloadFileAttachment(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)

                                    fileToSave = er.DownloadFileAttachment(dirToUploadDatabase, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)

                                    If vesselName = "meratus.karimata" Then
                                        fileToSave = er.DownloadFileAttachment("C:\yusuf\nlab\", pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                        att.Save(fileToSave)
                                    End If
                                End If

                            End If

                        End If

                    Else
                        'selain base64
                        If att.Filename <> "" Then
                            temp &= "-" & att.Filename & " "
                            Dim aaa As String = att.Filename.Substring(att.Filename.Length - 4, 3)
                            Dim fsf As String = att.ContentType & att.ContentDisposition & att.Body & att.AttachmentName
                            Dim xxx As String = att.Filename.Substring(1, 8)
                            'jika csv (ato mungkin file lain juga :D)
                            If att.Filename.Substring(att.Filename.Length - 4, 3) = "csv" Then
                                'If att.ContentType = "application/octet-stream" Or att.ContentType = "application/vnd.ms-excel" Or att.ContentType = "text/csv" Then
                                'add by yna 07.12.2015 for RMS
                                If att.Filename.Substring(1, 8) = "KW_METER" Or att.Filename.Substring(1, 8) = "kw_meter" Then
                                    Dim POP3Session As OSPOP3_Plus.Session
                                    POP3Session = New OSPOP3_Plus.Session
                                    dirToVesselArchive = dirDefaultRMS & dateNow & "\" & vesselName & "\"
                                    ''perkecualian benoa
                                    If pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.karimata" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.waigeo" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.manado" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.medan3" Then
                                        dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS\"
                                    Else
                                        dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS_mka\"
                                    End If
                                    Dim a As String = att.Filename.Substring(1, att.Filename.Length - 2)
                                    'disimpan 2 kali satunya untuk arsip satunya untuk diproses ke database kemudian dihapus (tempd_download folder)
                                    att.Body = POP3Session.QuotedPrintableDecode(att.Body())

                                    fileToSave = er.DownloadFileAttachment(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)
                                    'GetCsvData(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2), dirToVesselArchive & "tes" & pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))

                                    fileToSave = er.DownloadFileAttachment(dirToUploadDatabase, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    'att.Save(fileToSave)
                                    GetCsvData(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2), dirToUploadDatabase & pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))

                                    If pair.Value.Sender.Address.Split("@")(0) = "rms.meratus.karimata" Or pair.Value.Sender.Address.Split("@")(0) = "rms.meratus.waigeo" Then
                                        fileToSave = er.DownloadFileAttachment("C:\yusuf\nlab\", pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                        att.Save(fileToSave)
                                    End If
                                    'end of add by yna 07.12.2015 for RMS
                                    'add for encrypt data by yna 17.09.2018
                                ElseIf att.Filename.Substring(1, 9) = "eKW_METER" Or att.Filename.Substring(1, 9) = "ekw_meter" Then
                                    Dim POP3Session As OSPOP3_Plus.Session
                                    POP3Session = New OSPOP3_Plus.Session
                                    dirToVesselArchive = dirDefaultRMS & dateNow & "\" & vesselName & "\"
                                    'perkecualian benoa
                                    If pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.karimata" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.waigeo" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.manado" And pair.Value.Sender.Address.Split("@")(0) <> "rms.meratus.medan3" Then
                                        dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS\"
                                    Else
                                        dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS_mka\"
                                    End If
                                    dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS_encrypt\"
                                    Dim a As String = att.Filename.Substring(1, att.Filename.Length - 2)
                                    'disimpan 2 kali satunya untuk arsip satunya untuk diproses ke database kemudian dihapus (tempd_download folder)
                                    att.Body = POP3Session.QuotedPrintableDecode(att.Body())

                                    fileToSave = er.DownloadFileAttachment(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)
                                    'GetCsvData(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2), dirToVesselArchive & "tes" & pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))

                                    fileToSave = er.DownloadFileAttachment(dirToUploadDatabase, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    'att.Save(fileToSave)
                                    GetCsvData(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2), dirToUploadDatabase & pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))

                                    If pair.Value.Sender.Address.Split("@")(0) = "rms.meratus.karimata" Then
                                        fileToSave = er.DownloadFileAttachment("C:\yusuf\nlab\", pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                        att.Save(fileToSave)
                                    End If
                                    'end of add for encrypt data by yna 17.09.2018
                                ElseIf att.Filename.Substring(1, 3) = "RMS" Or att.Filename.Substring(1, 3) = "rms" Then
                                    Dim POP3Session As OSPOP3_Plus.Session
                                    POP3Session = New OSPOP3_Plus.Session
                                    dirToVesselArchive = dirDefaultRMS & dateNow & "\" & vesselName & "\"
                                    dirToUploadDatabase = dirDefaultRMS & dateNow & "\temp_downloadRMS_new\"
                                    Dim a As String = att.Filename.Substring(1, att.Filename.Length - 2)
                                    'disimpan 2 kali satunya untuk arsip satunya untuk diproses ke database kemudian dihapus (tempd_download folder)
                                    att.Body = POP3Session.QuotedPrintableDecode(att.Body())

                                    fileToSave = er.DownloadFileAttachment(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)
                                    'GetCsvData(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2), dirToVesselArchive & "tes" & pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))

                                    fileToSave = er.DownloadFileAttachment(dirToUploadDatabase, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    'att.Save(fileToSave)
                                    GetCsvData(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2), dirToUploadDatabase & pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    'end of add for encrypt data by yna 17.09.2018
                                Else
                                    Dim POP3Session As OSPOP3_Plus.Session
                                    POP3Session = New OSPOP3_Plus.Session

                                    'add by yna 25.07.2018 delete prefix rms.
                                    vesselName = vesselName.Replace("rms.", "")
                                    'end of add by yna 25.07.2018 delete prefix rms.

                                    dirToVesselArchive = dirDefault & dateNow & "\" & vesselName & "\"
                                    dirToUploadDatabase = dirDefault & dateNow & "\temp_download\"
                                    Dim a As String = att.Filename.Substring(1, att.Filename.Length - 2)
                                    'disimpan 2 kali satunya untuk arsip satunya untuk diproses ke database kemudian dihapus (tempd_download folder)
                                    att.Body = POP3Session.QuotedPrintableDecode(att.Body())

                                    fileToSave = er.DownloadFileAttachment(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    att.Save(fileToSave)
                                    'GetCsvData(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2), dirToVesselArchive & "tes" & pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))

                                    fileToSave = er.DownloadFileAttachment(dirToUploadDatabase, pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                    'att.Save(fileToSave)
                                    GetCsvData(dirToVesselArchive, pair.Value.UIDL & "_" & vesselName.Replace(".", "") & "_" & att.Filename.Substring(1, att.Filename.Length - 2), dirToUploadDatabase & pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))

                                    If vesselName = "meratus.karimata" Then
                                        fileToSave = er.DownloadFileAttachment("C:\yusuf\nlab\", pair.Value.UIDL & "_" & vesselName & "_" & att.Filename.Substring(1, att.Filename.Length - 2))
                                        att.Save(fileToSave)
                                    End If
                                End If

                            End If

                        End If

                    End If
                Next

                Console.WriteLine("Attach :" & temp)

                'Console.WriteLine("Body :" & pair.Value.Body)
                'Console.WriteLine("HTMLBody :" & pair.Value.HTMLBody)
                Console.WriteLine()
            End If
        Next

        er.CloseConnectionMailServer()
        'menjalankan proses update excel attachment ke dalam database menggunakan SSIS
        'RunBatFile()
        Console.WriteLine("_________________________________________________________________")

        'per 1 jam = 3600000 miliseconds
        'Thread.Sleep(3600000)
        'End While



        ' membaca data dari csv menjadi list atau dicionary dimana key adalah nama kolom
        'Try
        '    Dim list As New List(Of String())
        '    list = CSVReader.GetCSVData(fileToSave)

        '    Dim list2 As New List(Of Dictionary(Of String, String))
        '    list2 = CSVReader.GetCSVDatabyDictionary(fileToSave)
        'Catch ex As Exception
        '    Console.WriteLine(ex.Message)
        'End Try

        'insert data yang diperlukan ke dalam database (bisa menggunakan SP atau query standard),
        'memanfaatkan kelas UqeryHelper dan modifikasi sendiri untuk kasus tertentu


        'Console.Read()
    End Sub



    Private Sub InsertLog(ByVal message As String)
        Dim dateNow = DateTime.Now.Day & "-" & DateTime.Now.Month & "-" & DateTime.Now.Year
        If Not System.IO.Directory.Exists("Log\") Then
            System.IO.Directory.CreateDirectory("Log\")
        End If

        'log it
        Dim fs1 As FileStream = New FileStream("Log\LogMessage" & dateNow & ".txt", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)
        s1.WriteLine(message)
        s1.Close()
        fs1.Close()
    End Sub

    Private Sub RunBatFile()
        'Dim psi As New ProcessStartInfo("C:\febri\ServerInsertToDBase.bat")
        'psi.RedirectStandardError = True
        'psi.RedirectStandardOutput = True
        ''psi.CreateNoWindow = False
        'psi.WindowStyle = ProcessWindowStyle.Normal
        'psi.UseShellExecute = True
        'psi.Verb = "runas"

        'Dim process As Process = process.Start(psi)

        'process.WaitForExit()

        Try
            Dim procInfo As New ProcessStartInfo()
            procInfo.UseShellExecute = True
            procInfo.FileName = ("C:\GPS\Projek_Server_Febri\ServerInsertToDBase.bat")
            procInfo.WorkingDirectory = ""
            procInfo.Verb = "runas"
            Dim process As Process = process.Start(procInfo)

            process.WaitForExit()
        Catch ex As Exception
            Console.WriteLine(ex.Message.ToString())
        End Try
    End Sub

    Public Sub RemoveLogFile(ByVal daysAfter As Integer)
        Dim daysAfterFormatted = -1 * daysAfter
        Dim dateNow As Date = DateTime.Now.AddDays(daysAfterFormatted)
        Dim dateString = dateNow.Day & "-" & dateNow.Month & "-" & dateNow.Year

        Dim pathlog_att As String = My.Settings.path_log_attachment & dateString & "\"
        Dim pathlog_msg As String = My.Settings.path_log_message & dateString & "\"



        Try            
            If System.IO.Directory.Exists(pathlog_att) Then
                System.IO.Directory.Delete(pathlog_att, True)
            End If

            If System.IO.Directory.Exists(pathlog_msg) Then
                System.IO.Directory.Delete(pathlog_msg, True)
            End If

        Catch ex As Exception
        End Try


        'If My.Computer.FileSystem.FileExists(pathlog & "LogMessage" & dateString & ".txt") Then
        '    My.Computer.FileSystem.DeleteFile(pathlog & "LogMessage" & dateString & ".txt")
        'End If

    End Sub

    'add by yna 19.12.2016 utk yg tidak standard
    Public Sub GetCsvData(ByVal strFolderPath As String, ByVal strFileName As String, ByVal strFileNameDest As String)
        Try
            'CharacterSet=65001 will needed for UTF-8 settings
            Dim strConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFolderPath & ";Extended Properties='text;HDR=Yes;FMT=Delimited;CharacterSet=65001;'"
            Dim conn As New OleDbConnection(strConnString)
            Try
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT * FROM [" & strFileName & "]", conn)
                Dim da As New OleDbDataAdapter()

                da.SelectCommand = cmd
                Dim ds As New DataSet()

                da.Fill(ds)
                da.Dispose()
                DataTable2CSV(ds.Tables(0), strFileNameDest, ",")

                'Return ds.Tables(0)
            Catch x As Exception
                'Return Nothing
            Finally
                conn.Close()
            End Try
        Catch ex As Exception
        End Try
    End Sub

    Public Sub DataTable2CSV(ByVal table As DataTable, ByVal filename As String, _
ByVal sepChar As String)
        Dim writer As System.IO.StreamWriter
        Try
            writer = New System.IO.StreamWriter(filename)

            ' first write a line with the columns name
            Dim sep As String = ""
            Dim builder As New System.Text.StringBuilder
            For Each col As DataColumn In table.Columns
                builder.Append(sep).Append(col.ColumnName)
                sep = sepChar
            Next
            writer.WriteLine(builder.ToString())

            ' then write all the rows
            Dim cnt As Integer
            cnt = 0
            For Each row As DataRow In table.Rows
                sep = ""
                builder = New System.Text.StringBuilder
                cnt = 0
                For Each col As DataColumn In table.Columns
                    builder.Append(sep).Append(row(col.ColumnName))
                    sep = sepChar
                    If row(0) = "" Then
                        cnt = cnt + 1
                    End If
                Next
                'yg kosong2 di skip aja
                If cnt = 0 Then
                    writer.WriteLine(builder.ToString())
                End If
            Next
        Finally
            If Not writer Is Nothing Then writer.Close()
        End Try
    End Sub
    'end of add by yna 19.12.2016 utk yg tidak standard

    'function for decrypt by yna 17.09.2018
    Public Sub setDecryptFile()
        'writer.CreateCSVHeader("Z:\FLOWMETER" & "-" & datenow.Hour & "-" & datenow.Minute & ".csv", header)
        'list the names of all files in the specified directory
        Try
            Dim di As New IO.DirectoryInfo("C:\test\encrypt\")
            Dim infoFile As IO.FileInfo() = di.GetFiles()


            Dim info As IO.FileInfo
            Dim writer As New CSVWriter()
            Dim header As New List(Of String)
            Dim datenow As Date = DateTime.Now
            header.Add(" TIME")
            header.Add(" WATT_GEN1")
            header.Add(" VLN_GEN1")
            header.Add(" WATT_GEN2")
            header.Add(" VLN_GEN2")
            header.Add(" WATT_GEN3")
            header.Add(" VLN_GEN3")
            header.Add(" WATT_GEN4")
            header.Add(" VLN_GEN4")
            header.Add(" TEMP_1")
            header.Add(" TEMP_2")
            header.Add(" TEMP_3")
            header.Add(" TEMP_4")
            header.Add(" VOLUME")
            header.Add(" FLAG")

            writer.CreateDirectory("C:\test") '

            For Each info In infoFile
                Console.WriteLine(info.FullName)

                Dim tempInfoName As String() = info.Name.Split("-")

                Dim getData As List(Of String()) = CSVReader.GetCSVData(info.FullName)

                'writer.CreateCSVHeader("Z:\flowmeter\" & temp(3).Split(".")(0) & "-" & datenow.Hour & "-" & datenow.Minute & ".csv", header)
                'writer.CreateCSVHeader("C:\test\decrypt\" & tempInfoName(3).Split(".")(0) & "-" & tempInfoName(1) & "-" & tempInfoName(2) & ".csv", header)
                writer.CreateCSVHeader("C:\test\decrypt\" & tempInfoName(0) & "-" & tempInfoName(1) & "-" & tempInfoName(2) & "", header)

                For Each iter As String() In getData
                    For Each Data As String In iter
                        Dim result As String = LightWeightEncrypt(Data, 80)
                        Dim content As New List(Of String)
                        'Console.WriteLine(Data & " = " & result)
                        'content.Add(Data)
                        content.Add(result)
                        'data akhir

                        writer.CreateCSVAppend("C:\test\decrypt\" & tempInfoName(0) & "-" & tempInfoName(1) & "-" & tempInfoName(2) & "", content)
                    Next
                Next
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        'Console.ReadLine()

    End Sub

    Public Function LightWeightEncrypt(ByVal Text As String, ByVal EncryptKey As Integer) As String

        Dim strTemp As String = ""
        Dim intCounter As Integer
        For intCounter = 1 To Len(Text)
            strTemp$ = strTemp$ + Chr(Asc(Mid(Text, intCounter, 1)) Xor EncryptKey)
        Next intCounter
        Return strTemp
    End Function
    'end of function for decrypt by yna 17.09.2018

End Module
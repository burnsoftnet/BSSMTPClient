Imports System.Net.Mail
Imports System.Net
Namespace BurnSoft
    Public Class BSMail
        ''' <summary>
        ''' This sub will send an email to a person or group of people in HTML Format
        ''' </summary>
        ''' <param name="strTo"></param>
        ''' <param name="strFrom"></param>
        ''' <param name="strSubject"></param>
        ''' <param name="strMessage"></param>
        Public Sub SendMail(ByVal strTo As String, ByVal strFrom As String, ByVal strSubject As String, ByVal strMessage As String, Optional ByVal CC As String = "", Optional ByVal BCC As String = "")
            If Len(MAIL_PASSWORD) > 0 Then
                strFrom = MAIL_USER
            End If
            Dim strSendFrom As MailAddress = New MailAddress(strFrom)
            Dim Message As MailMessage = New MailMessage
            Dim i As Integer = 0
            Dim strSplit As Array = Split(strTo, ",")
            Dim intBound As Integer = UBound(strSplit)
            Dim Client As New System.Net.Mail.SmtpClient
            Dim ErrorMsg As String = ""
            Message.From = strSendFrom

            If intBound <> 0 Then
                For i = 0 To intBound - 1
                    Message.To.Add(strSplit(i))
                Next
            Else
                Message.To.Add(strTo)
            End If

            If Len(CC) > 0 Then
                strSplit = Split(CC, ",")
                intBound = UBound(strSplit)
                If intBound <> 0 Then
                    For i = 0 To intBound - 1
                        Message.CC.Add(strSplit(i))
                    Next
                Else
                    Message.CC.Add(strTo)
                End If
            End If

            If Len(BCC) > 0 Then
                strSplit = Split(BCC, ",")
                intBound = UBound(strSplit)
                If intBound <> 0 Then
                    For i = 0 To intBound - 1
                        Message.Bcc.Add(strSplit(i))
                    Next
                Else
                    Message.Bcc.Add(strTo)
                End If
            End If

            Try
                Message.IsBodyHtml = MAIL_USEHTML
                Message.Subject = strSubject
                Message.Body = strMessage

                If Len(MAIL_ATTACHFILES) > 0 Then
                    strSplit = Split(MAIL_ATTACHFILES, "'")
                    intBound = UBound(strSplit)
                    Dim myFile As Attachment
                    If intBound <> 0 Then
                        For i = 0 To intBound - 1
                            If FileExists(strSplit(i)) Then
                                myFile = New Attachment(strSplit(i))
                                Message.Attachments.Add(myFile)
                            End If
                        Next
                    Else
                        If FileExists(MAIL_ATTACHFILES) Then
                            myFile = New Attachment(MAIL_ATTACHFILES)
                            Message.Attachments.Add(myFile)
                        End If
                    End If
                End If


                Client.Host = MAIL_SERVER
                Client.Port = MAIL_PORT
                Client.Timeout = MAIL_TIMEOUT
                Client.EnableSsl = MAIL_USETLS

                If Len(MAIL_PASSWORD) > 0 Then Client.Credentials = New Net.NetworkCredential(MAIL_USER, MAIL_PASSWORD)

                'If PROXY_USE Then
                ' Dim oProxy As New WebProxy(PROXY_SERVER, PROXY_PORT)
                ' oProxy.Credentials() = New NetworkCredential(PROXY_UID, PROXY_PWD)
                'Client.Credentials = oProxy
                'End If

                Client.Send(Message)
                Catch ex As Exception
                    ErrorMsg = "ERROR: " & ex.Message.ToString
                Finally
                    Message.Dispose()
                    Message = Nothing
                    Client = Nothing
                    If Len(ErrorMsg) > 0 Then
                        Console.WriteLine(ErrorMsg)
                        If LOG_TO_FILE Then Call LogFile(LOG_PATH & "\" & LOG_NAME, ErrorMsg)
                    Else
                        Console.WriteLine("Message Sent to: " & strTo)
                        If LOG_TO_FILE Then Call LogFile(LOG_PATH & "\" & LOG_NAME, "Message Sent to: " & strTo)
                    End If
                End Try
        End Sub
    End Class
End Namespace
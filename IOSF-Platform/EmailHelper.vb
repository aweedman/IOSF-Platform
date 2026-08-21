Imports MailKit.Net.Smtp
Imports MailKit.Security
Imports MimeKit

''' <summary>
''' Sends email via SMTP (port 587, STARTTLS) using MailKit.
'''
''' The SMTP server and From address are read from Config rather than hardcoded, so this
''' repository never contains real credentials. Required Config rows (alongside
''' "Email Pass"):
'''   Name = "Email SMTP Server", Low = Server
'''   Name = "Email From Address", Low = Email Address
''' </summary>
Public Module EmailHelper

    Private Const SmtpPort As Integer = 587

    ''' <summary>Sends an email. bcc/replyTo/attachmentPath are optional - pass Nothing/empty to omit any of them.</summary>
    Public Sub SendEmail(toAddress As String, subject As String, body As String,
                          Optional bcc As String = Nothing,
                          Optional replyTo As String = Nothing,
                          Optional attachmentPath As String = Nothing)
        Dim smtpServer = ConfigHelper.GetConfigValue("Email SMTP Server")
        Dim fromAddress = ConfigHelper.GetConfigValue("Email From Address")
        Dim password = ConfigHelper.GetConfigValue("Email Pass")

        Dim message As New MimeMessage()
        message.From.Add(MailboxAddress.Parse(fromAddress))
        message.To.Add(MailboxAddress.Parse(toAddress))
        message.Subject = subject

        If Not String.IsNullOrEmpty(bcc) Then
            message.Bcc.Add(MailboxAddress.Parse(bcc))
        End If

        If Not String.IsNullOrEmpty(replyTo) Then
            message.ReplyTo.Add(MailboxAddress.Parse(replyTo))
        End If

        Dim builder As New BodyBuilder With {.TextBody = body}
        If Not String.IsNullOrEmpty(attachmentPath) Then
            builder.Attachments.Add(attachmentPath)
        End If
        message.Body = builder.ToMessageBody()

        Using client As New SmtpClient()
            client.Connect(smtpServer, SmtpPort, SecureSocketOptions.StartTls)
            client.Authenticate(fromAddress, password)
            client.Send(message)
            client.Disconnect(True)
        End Using
    End Sub

    ''' <summary>Sends an error notification to the address configured under "Email Error User".</summary>
    Public Sub EmailError(message As String)
        Dim toUser = ConfigHelper.GetConfigValue("Email Error User")
        SendEmail(toUser, "IOSF-Platform Error", message)
    End Sub

End Module
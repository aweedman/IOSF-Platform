Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Replaces Command65_Click's five instructional MsgBoxes plus three InputBox prompts
''' with one dialog. Instructional text is copied verbatim from the original rather than
''' paraphrased, since it's a step-by-step guide to an external site (developer.remotelock.com)
''' and small wording differences could genuinely confuse someone following along.
''' </summary>
Public Class RemoteLockAuthDialog
    Inherits Form

    Private txtClientId As TextBox
    Private txtClientSecret As TextBox
    Private txtCode As TextBox

    Public ReadOnly Property ClientId As String
        Get
            Return txtClientId.Text
        End Get
    End Property

    Public ReadOnly Property ClientSecret As String
        Get
            Return txtClientSecret.Text
        End Get
    End Property

    Public ReadOnly Property Code As String
        Get
            Return txtCode.Text
        End Get
    End Property

    Public Sub New()
        Text = "RemoteLock Refresh Token"
        ClientSize = New Size(460, 360)
        StartPosition = FormStartPosition.CenterParent
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        MinimizeBox = False

        Dim instructions As New Label With {
            .Text =
                "Go to URL developer.remotelock.com and sign in" & vbCrLf &
                "'Destroy' current token" & vbCrLf &
                "Click 'New OAuth Application' and Enter a Name" & vbCrLf &
                "Redirect URI is 'urn:ietf:wg:oauth:2.0:oob' and click 'Submit'" & vbCrLf &
                "Click 'Authorize', log in, and click 'Authorize'",
            .Location = New Point(20, 20),
            .Width = 420,
            .Height = 100
        }

        Dim lblClientId As New Label With {.Text = "Client ID", .Location = New Point(20, 130), .AutoSize = True}
        txtClientId = New TextBox With {.Location = New Point(20, 150), .Width = 420}

        Dim lblClientSecret As New Label With {.Text = "Client Secret", .Location = New Point(20, 180), .AutoSize = True}
        txtClientSecret = New TextBox With {.Location = New Point(20, 200), .Width = 420}

        Dim lblCode As New Label With {.Text = "Code", .Location = New Point(20, 230), .AutoSize = True}
        txtCode = New TextBox With {.Location = New Point(20, 250), .Width = 420}

        Dim btnOk As New Button With {.Text = "OK", .Location = New Point(275, 300), .Width = 75, .DialogResult = DialogResult.OK}
        Dim btnCancel As New Button With {.Text = "Cancel", .Location = New Point(365, 300), .Width = 75, .DialogResult = DialogResult.Cancel}

        Controls.AddRange({instructions, lblClientId, txtClientId, lblClientSecret, txtClientSecret, lblCode, txtCode, btnOk, btnCancel})
        AcceptButton = btnOk
        CancelButton = btnCancel
    End Sub

End Class
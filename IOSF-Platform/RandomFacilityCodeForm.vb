Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms

''' <summary>
''' Generates a random, unused facility code:
'''  - Generate Code: picks a random 5-digit number (00000-99999), checks
'''    Used_Fac_Codes.Code for a collision, and keeps picking until an unused one is
'''    found, then displays it.
'''  - Use Code: inserts the currently-shown code into Used_Fac_Codes, reserving it.
'''
''' Collision-checking fetches all currently-used codes into memory once up front, then
''' loops randomly in-memory until a non-colliding 5-digit number is found - not a
''' separate database round-trip per random attempt.
''' </summary>
Public Class RandomFacilityCodeForm
    Inherits Form

    Private ReadOnly rng As New Random()
    Private currentCode As String

    Private codeBox As TextBox
    Private copyButton As Button
    Private generateButton As Button
    Private useButton As Button

    Public Sub New()
        Text = "Random Facility Code"
        Width = 500
        Height = 300
        StartPosition = FormStartPosition.CenterScreen

        ' TextBox, not Label - Labels don't support text selection/Ctrl+C in WinForms.
        ' ReadOnly keeps it non-editable while still allowing selection/copy.
        codeBox = New TextBox With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .ReadOnly = True,
            .TextAlign = HorizontalAlignment.Center,
            .Font = New Font("Segoe UI", 28, FontStyle.Bold),
            .BorderStyle = BorderStyle.None,
            .BackColor = SystemColors.Control,
            .Text = ""
        }

        copyButton = New Button With {.Text = "Copy to Clipboard", .Dock = DockStyle.Top, .Height = 40, .Enabled = False}
        AddHandler copyButton.Click, AddressOf CopyCodeClicked

        generateButton = New Button With {.Text = "Generate Code", .Dock = DockStyle.Top, .Height = 60}
        AddHandler generateButton.Click, AddressOf GenerateCodeClicked

        useButton = New Button With {.Text = "Use Code", .Dock = DockStyle.Top, .Height = 60, .Enabled = False}
        AddHandler useButton.Click, AddressOf UseCodeClicked

        ' Stacked Top controls - order added is reversed visual order, last-added ends
        ' up outermost/topmost.
        Controls.Add(useButton)
        Controls.Add(generateButton)
        Controls.Add(copyButton)
        Controls.Add(codeBox)
    End Sub

    Private Sub CopyCodeClicked(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(currentCode) Then Return
        Clipboard.SetText(currentCode)
    End Sub

    Private Sub GenerateCodeClicked(sender As Object, e As EventArgs)
        Try
            Dim usedCodes = FetchUsedCodes()

            Const maxAttempts As Integer = 10000 ' safety cap - astronomically unlikely to ever be hit given up to 100,000 possible 5-digit codes, but avoids a true infinite loop in the pathological case where all of them are somehow used
            Dim attempts = 0
            Dim candidate As String

            Do
                candidate = rng.Next(0, 100000).ToString("D5") ' zero-padded to always be 5 digits
                attempts += 1
                If attempts > maxAttempts Then
                    MessageBox.Show(Me, "Could not find an unused code after many attempts - this would mean nearly all 100,000 possible codes are already in use.", "Random Facility Code", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            Loop While usedCodes.Contains(candidate)

            currentCode = candidate
            codeBox.Text = currentCode
            useButton.Enabled = True
            copyButton.Enabled = True
        Catch ex As Exception
            MessageBox.Show(Me, $"Error generating code: {ex.Message}", "Random Facility Code", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function FetchUsedCodes() As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.Ordinal)
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand("SELECT Code FROM Used_Fac_Codes", conn)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        If Not reader.IsDBNull(0) Then result.Add(reader.GetString(0))
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Sub UseCodeClicked(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(currentCode) Then Return
        Try
            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                Using cmd As New SqlCommand("INSERT INTO Used_Fac_Codes (Code) VALUES (@Code)", conn)
                    cmd.Parameters.AddWithValue("@Code", currentCode)
                    conn.Open()
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show(Me, "Code Reserved - Please Make Note of It", "Random Facility Code", MessageBoxButtons.OK, MessageBoxIcon.Information)
            useButton.Enabled = False ' once used, a new code must be generated before using again, rather than allowing a double-insert of the same code
        Catch ex As Exception
            MessageBox.Show(Me, $"Error reserving code: {ex.Message}", "Random Facility Code", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
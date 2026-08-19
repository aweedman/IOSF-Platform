Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms

''' <summary>
''' Ports the "Random_Facility_Code" PowerApp. The app itself called two Power Automate
''' flows (RandomFacilityCode2, EnterUsedFacilityCode2) rather than SQL Server directly -
''' those flows' internal logic isn't visible in the .msapp package (Power Automate flows
''' are separate, externally-hosted definitions), so this is built directly from Al's own
''' description of what they do, not from inspecting the flows themselves:
'''  - Generate Code: pick a random 5-digit number (00000-99999), check Used_Fac_Codes.Code
'''    for a collision, keep picking until an unused one is found, then show it.
'''  - Use Code: insert the currently-shown code into Used_Fac_Codes, confirming it's now reserved.
'''
''' Collision-checking is done via ONE query fetching all currently-used codes into memory
''' up front, then looping randomly in-memory until a non-colliding 5-digit number is
''' found - not a separate database round-trip per random attempt. This should be
''' functionally equivalent to whatever the original flow did (same end result: an unused
''' code), just more efficient given up to 100,000 possible codes.
'''
''' Table/column name (Used_Fac_Codes.Code) confirmed directly by Al, not guessed.
''' </summary>
Public Class RandomFacilityCodeForm
    Inherits Form

    Private ReadOnly rng As New Random()
    Private currentCode As String

    Private codeLabel As Label
    Private generateButton As Button
    Private useButton As Button

    Public Sub New()
        Text = "Random Facility Code"
        Width = 500
        Height = 260
        StartPosition = FormStartPosition.CenterScreen

        codeLabel = New Label With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .TextAlign = ContentAlignment.MiddleCenter,
            .Font = New Font("Segoe UI", 28, FontStyle.Bold),
            .Text = ""
        }

        generateButton = New Button With {.Text = "Generate Code", .Dock = DockStyle.Top, .Height = 60}
        AddHandler generateButton.Click, AddressOf GenerateCodeClicked

        useButton = New Button With {.Text = "Use Code", .Dock = DockStyle.Top, .Height = 60, .Enabled = False}
        AddHandler useButton.Click, AddressOf UseCodeClicked

        ' Fill-first-then-edges rule (see other forms' remarks for why this order
        ' matters) - none of these are Fill here, just stacked Top controls, so order is
        ' simply reversed visual order: last-added ends up outermost/topmost.
        Controls.Add(useButton)
        Controls.Add(generateButton)
        Controls.Add(codeLabel)
    End Sub

    Private Sub GenerateCodeClicked(sender As Object, e As EventArgs)
        Try
            Dim usedCodes = FetchUsedCodes()

            Const maxAttempts As Integer = 10000 ' safety cap - astronomically unlikely to ever be hit given up to 100,000 possible 5-digit codes, but avoids a true infinite loop in the pathological case where all of them are somehow used
            Dim attempts = 0
            Dim candidate As String

            Do
                candidate = rng.Next(0, 100000).ToString("D5") ' zero-padded to always be 5 digits, matching "00000 to 99999"
                attempts += 1
                If attempts > maxAttempts Then
                    MessageBox.Show(Me, "Could not find an unused code after many attempts - this would mean nearly all 100,000 possible codes are already in use.", "Random Facility Code", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            Loop While usedCodes.Contains(candidate)

            currentCode = candidate
            codeLabel.Text = currentCode
            useButton.Enabled = True
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
            useButton.Enabled = False ' matches the original's own implicit flow - once used, generate a new one before using again, rather than allowing a double-insert of the same code
        Catch ex As Exception
            MessageBox.Show(Me, $"Error reserving code: {ex.Message}", "Random Facility Code", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
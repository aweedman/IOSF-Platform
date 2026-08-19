Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command23_Click() ("180.1 - Spheremail Charges to
''' DB").
'''
''' Reuses SphereMailAuth.GetTokenAsync() rather than reimplementing authentication - it
''' already fixes a real latent bug present in the original: the VBA never checked whether
''' the login call succeeded before using the returned token, so a failed auth would
''' silently proceed with an empty/uninitialized token rather than aborting (see
''' SphereMailStorageJob's remarks for the same issue, fixed the same way there).
'''
''' "Burlingame" removed: the original looped "For i = 1 To 1" with a dead Else branch for
''' a second location ("Burlingame") that could never be reached - the loop only ever runs
''' once. Burlingame was a location decommissioned before this porting effort began (same
''' pattern already established in SphereMailCustomersJob/SphereMailStorageJob), so this is
''' simplified to a plain "San Francisco" constant instead of reproducing a loop structure
''' that no longer does anything.
'''
''' Manual sequential Id: the original computes MAX(Id) once via DMax and increments it
''' itself for each inserted row, rather than an IDENTITY column - preserved exactly
''' (ISNULL-guarded here in case the table is ever empty, which the original's DMax-into-a-
''' typed-Long variable would have crashed on, though this is an extreme edge case
''' unlikely to matter in practice).
'''
''' Description filter preserved exactly: only "Scan", "Mail Item Picture",
''' "Mail Item Pickup", and "joint_account" items are inserted - everything else is
''' silently skipped, matching the original's own filter.
'''
''' Email/First_Name/Last_Name are always stored blank, matching the original exactly
''' (presumably filled in by a separate, not-yet-ported process).
'''
''' ERROR HANDLING - deliberately NOT matching Call Counts/Variable Charges' atomic
''' transaction pattern, since Al did not ask for that here (only for Variable Charges,
''' explicitly). This keeps the original's own per-row On Error Resume Next behavior: a
''' failure on one charge item is logged and does not stop the rest or roll back what
''' already succeeded. Worth flagging in case an atomic version is actually preferred here
''' too, once tested.
'''
''' Table name NOT independently verified: Spheremail_Charges_SQL -> assumed real name
''' Spheremail_Charges (the simple-strip convention that's held for most tables in this
''' port, but not confirmed against a tbldefs descriptor).
'''
''' No pagination: the original requests a single page with limit=10000 and never loops -
''' preserved exactly, since there's no evidence this specific endpoint supports/returns
''' pagination metadata the way the HostedSuite APIs do.
''' </summary>
Public Module SpheremailChargesToDbJob

    Private Const ApiBaseUrl As String = "https://api.spheremail.co/v1/admin"
    Private Const LocationName As String = "San Francisco" ' see class remarks re: Burlingame
    Private ReadOnly IncludedDescriptions As String() = {"Scan", "Mail Item Picture", "Mail Item Pickup", "joint_account"}

    Public Async Function RunAsync(startDate As Date, endDate As Date) As Task(Of Integer)
        Dim errorCount = 0
        Dim charges As SphereMailChargesResponse

        Try
            Dim token = Await SphereMailAuth.GetTokenAsync()
            Dim headers = New Dictionary(Of String, String) From {{"Authorization", token}}
            Dim queryParams = New Dictionary(Of String, String) From {
                {"from", startDate.ToString("yyyy-MM-dd")},
                {"to", endDate.ToString("yyyy-MM-dd")},
                {"limit", "10000"}
            }

            Dim response = Await ApiClient.GetAsync($"{ApiBaseUrl}/reports/charges/detail", queryParams, headers, timeoutSeconds:=15)
            response.EnsureSuccess()
            charges = response.DataAs(Of SphereMailChargesResponse)()
        Catch ex As Exception
            ErrorLogHelper.LogError("Spheremail Charges to DB", $"API Call Error: {ex.Message}")
            Return 1
        End Try

        errorCount += ApplySpheremailCharges(startDate, endDate, charges)
        Return errorCount
    End Function

    Private Function ApplySpheremailCharges(startDate As Date, endDate As Date, charges As SphereMailChargesResponse) As Integer
        Dim errorCount = 0

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            conn.Open()

            Try
                Using deleteCmd As New SqlCommand("DELETE FROM Spheremail_Charges WHERE Txn_Date BETWEEN @StartDate AND @EndDate", conn)
                    deleteCmd.Parameters.AddWithValue("@StartDate", startDate)
                    deleteCmd.Parameters.AddWithValue("@EndDate", endDate)
                    deleteCmd.ExecuteNonQuery()
                End Using
            Catch ex As Exception
                ErrorLogHelper.LogError("Spheremail Charges to DB", $"SQL error in: DELETE FROM Spheremail_Charges - {ex.Message}")
                errorCount += 1
            End Try

            Dim nextId As Integer
            Using maxCmd As New SqlCommand("SELECT ISNULL(MAX(Id), 0) FROM Spheremail_Charges", conn)
                nextId = CInt(maxCmd.ExecuteScalar())
            End Using

            If charges.Charges IsNot Nothing Then
                For Each group In charges.Charges
                    If group.Items Is Nothing Then Continue For

                    For Each item In group.Items
                        If Not IncludedDescriptions.Contains(item.Description) Then Continue For

                        nextId += 1

                        Try
                            Const insertSql As String =
                                "INSERT INTO Spheremail_Charges (Id, Txn_Date, Description, Email, First_Name, Last_Name, Mail_Box, Location) " &
                                "VALUES (@Id, @TxnDate, @Description, '', '', '', @MailBox, @Location)"

                            Using cmd As New SqlCommand(insertSql, conn)
                                cmd.Parameters.AddWithValue("@Id", nextId)
                                cmd.Parameters.AddWithValue("@TxnDate", DateTime.Parse(item.Date))
                                cmd.Parameters.AddWithValue("@Description", item.Description)
                                cmd.Parameters.AddWithValue("@MailBox", group.PmbNumber)
                                cmd.Parameters.AddWithValue("@Location", LocationName)
                                cmd.ExecuteNonQuery()
                            End Using
                        Catch ex As Exception
                            ErrorLogHelper.LogError("Spheremail Charges to DB", $"SQL error inserting charge Id={nextId}, PMB={group.PmbNumber}: {ex.Message}")
                            errorCount += 1
                        End Try
                    Next
                Next
            End If
        End Using

        Return errorCount
    End Function

End Module
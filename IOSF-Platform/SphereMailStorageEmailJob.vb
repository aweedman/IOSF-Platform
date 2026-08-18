Imports Microsoft.Data.SqlClient
Imports System.IO

''' <summary>
''' Direct port of Landing Page.cls: Command26_Click() ("Spheremail Storage Emails" button).
''' Runs Spheremail_Storage(30), then either emails staff one aggregate PDF, or emails
''' each affected customer their own filtered PDF.
'''
''' Changes from the VBA original:
'''  - Spheremail_Storage_Temp was a local Access table (see SphereMailStorageJob remarks)
'''    used only for staging rows for the bound report. Filtering by location/mailbox
'''    number now happens directly against the in-memory List(Of SphereMailStorageRow)
'''    SphereMailStorageJob returns - no DB round-trip.
'''  - The dead "Burlingame" loop branch is dropped, per the original's own comment
'''    ("Burlingame would be 2 if API was active") confirming it's intentionally inactive.
'''  - Mode selection (staff summary vs. individual customers) and the interactive
'''    "Run for Individual Customers?" prompt are pulled OUT of this function into a
'''    Mode parameter - the caller is responsible for deciding the mode and showing any
'''    prompt. Same for the "Process Complete" MsgBox: this function returns the error
'''    count instead of deciding whether/how to display it.
'''  - Hardcoded "C:\AccessTemp\Mail Storage.pdf" replaced with Path.GetTempPath().
'''  - DCount/DLookup calls are now parameterized queries, or plain LINQ against the
'''    in-memory row list where the original was counting/filtering staged rows.
'''  - Per-customer errors (email lookup failure) are caught per-iteration so one bad
'''    customer doesn't abort the whole loop.
'''  - The staff notification address ("sffidi.virtual@intelligentoffice.com" in the
'''    original, used as the staff-summary recipient and as BCC/ReplyTo on customer
'''    emails) is now read from Config instead of hardcoded, so it's not exposed in a
'''    public repo - add a row: Name = "Spheremail Staff Notification Email".
''' </summary>
Public Module SphereMailStorageEmailJob

    Public Enum Mode
        StaffSummary        ' original: All = "7"
        IndividualCustomers ' original: All = "2" / anything else
    End Enum

    Private Const Location As String = "San Francisco"
    Private ReadOnly TempPdfPath As String = Path.Combine(Path.GetTempPath(), "Mail Storage.pdf")

    ''' <summary>
    ''' Runs the full flow and returns the number of errors logged, so the caller can
    ''' decide how to surface completion (MsgBox if interactive, EmailError if batch).
    ''' </summary>
    Public Async Function RunAsync(mode As Mode) As Task(Of Integer)
        Dim errorCount = 0

        Dim result = Await SphereMailStorageJob.RunAsync(30) ' original: Call Spheremail_Storage(30)
        Dim rows = result.Rows
        errorCount += result.ErrorCount

        If mode = Mode.StaffSummary Then
            errorCount += Await RunStaffSummaryAsync(rows)
        Else
            errorCount += Await RunIndividualCustomersAsync(rows)
        End If

        Return errorCount
    End Function

    Private Async Function RunStaffSummaryAsync(rows As List(Of SphereMailStorageRow)) As Task(Of Integer)
        Dim locationRows = rows.Where(Function(r) r.Location = Location).ToList()
        If locationRows.Count = 0 Then Return 0

        Try
            Await ReportGenerator.GenerateSphereMailStoragePdfAsync(locationRows, Location, TempPdfPath)

            EmailHelper.SendEmail(
                toAddress:=ConfigHelper.GetConfigValue("Spheremail Staff Notification Email"),
                subject:="Spheremail Storage Over 30 Days",
                body:="Please see the attachment. Customers will receive their extended mail storage notification automatically tomorrow. " &
                      "Please update Spheremail as needed to ensure accuracy." & vbCrLf & vbCrLf & "Al",
                attachmentPath:=TempPdfPath)

            Return 0
        Catch ex As Exception
            ErrorLogHelper.LogError("Spheremail Storage Emails", $"Error generating/sending staff summary: {ex.Message}")
            Return 1
        Finally
            DeleteTempPdfIfExists()
        End Try
    End Function

    Private Async Function RunIndividualCustomersAsync(rows As List(Of SphereMailStorageRow)) As Task(Of Integer)
        Dim errorCount = 0
        Dim mailboxNumbers = rows.
            Where(Function(r) r.Location = Location).
            Select(Function(r) r.PrivateMailboxNumber).
            Distinct().
            ToList()

        Dim staffEmail = ConfigHelper.GetConfigValue("Spheremail Staff Notification Email")

        For Each mailboxNumber In mailboxNumbers
            Try
                Dim email = GetCustomerEmail(mailboxNumber, Location)

                If String.IsNullOrEmpty(email) Then
                    ErrorLogHelper.LogError("Spheremail Storage Emails", $"Unable to determine email address for PMB {mailboxNumber}")
                    errorCount += 1
                    Continue For
                End If

                Dim customerRows = rows.
                    Where(Function(r) r.Location = Location AndAlso r.PrivateMailboxNumber = mailboxNumber).
                    ToList()

                Await ReportGenerator.GenerateSphereMailStoragePdfAsync(customerRows, Location, TempPdfPath)

                EmailHelper.SendEmail(
                    toAddress:=email,
                    subject:="Digital Address Storage Over 30 Days",
                    body:="Hello. Please see the attachment for your mail items that were received over 30 days ago. " &
                          "We have limited space to store mail. Kindly log into www.intelligentofficesf.com and either " &
                          "mark these items for shredding or mail forward. You may also come in and pick up. Please be " &
                          "advised that extended storage incurs a $1/day charge. Thank you." & vbCrLf & vbCrLf &
                          "Intelligent Office" & vbCrLf &
                          "100 Pine Street, Suite 1250 | San Francisco, CA 94111" & vbCrLf &
                          "Office: 415-745-3300 | Fax: 415-745-3301" & vbCrLf &
                          "sf.intelligentoffice.com",
                    bcc:=staffEmail,
                    replyTo:=staffEmail,
                    attachmentPath:=TempPdfPath)

                DeleteTempPdfIfExists()
                Await Task.Delay(1000) ' original: Sleep 1000

            Catch ex As Exception
                ErrorLogHelper.LogError("Spheremail Storage Emails", $"Error processing PMB {mailboxNumber}: {ex.Message}")
                errorCount += 1
            End Try
        Next

        Return errorCount
    End Function

    Private Sub DeleteTempPdfIfExists()
        If File.Exists(TempPdfPath) Then File.Delete(TempPdfPath)
    End Sub

    Private Function GetCustomerEmail(mailboxNumber As String, location As String) As String
        Const sql As String =
            "SELECT Email FROM Spheremail_Customers WHERE Mail_Box = @MailBox AND Location = @Location"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@MailBox", mailboxNumber)
                cmd.Parameters.AddWithValue("@Location", location)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                Return If(result Is Nothing OrElse result Is DBNull.Value, String.Empty, result.ToString())
            End Using
        End Using
    End Function

End Module
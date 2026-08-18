Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Spheremail_Storage(Days).
''' Finds stored mail items older than Days, resolves each to a customer name via
''' Customer_Ops -> Customer_QB, and returns them as an in-memory list.
'''
''' Changes from the VBA original:
'''  - Spheremail_Storage_Temp turns out to be a genuine LOCAL Access table (confirmed via
'''    its tbldefs export - no linked-table .json descriptor, unlike every _SQL-suffixed
'''    table). Its only purpose was staging rows for a bound report. .NET reports don't
'''    need a bound recordsource, so this now just returns List(Of SphereMailStorageRow)
'''    directly - no DB round-trip for staging at all.
'''  - The DLookup chain (Customer_Ops -> "Customer_Sync_From_QB") targets a table whose
'''    ALIAS is "Customer_Sync_From_QB_SQL" but whose actual SQL Server table name is
'''    "Customer_QB" (confirmed via its tbldefs .json - SourceTableName: dbo.Customer_QB).
'''    Naive suffix-stripping would have gotten this wrong; verified against the real
'''    linked-table descriptor instead. Note there's a SEPARATE "Customer_QB" Access alias
'''    that's QODBC-linked straight to QuickBooks' live Customer table - not the same thing
'''    as this one, which is SQL Server's own synced copy.
'''  - The "Days = 999 triggers an interactive DatePicker prompt" branch is NOT ported here
'''    - that's UI logic and belongs in whatever click handler calls this. This function
'''    now always takes a concrete Days value.
'''  - On SphereMail auth failure, the original's Resume Next handler had a latent bug
'''    (continues with an empty token instead of stopping) - not reproduced; this aborts
'''    cleanly via SphereMailAuth.GetTokenAsync's EnsureSuccess.
'''  - DLookup calls are now parameterized queries.
''' </summary>
Public Module SphereMailStorageJob

    Private Const AdminBaseUrl As String = "https://api.spheremail.co/v1/admin"
    Private Const Location As String = "San Francisco" ' matches SphereMailCustomersJob

    ''' <summary>
    ''' REAL BUG FIXED: this used to return only List(Of SphereMailStorageRow), with any
    ''' failure (auth, customer sync, mail-item fetch) caught internally and swallowed into
    ''' an empty list - indistinguishable, from the caller's perspective, from "genuinely
    ''' nothing to report". SphereMailStorageEmailJob.RunAsync treated an empty list as
    ''' success and returned 0 errors, so a real failure (confirmed: a JSON deserialization
    ''' error on the very first customer-sync call) got logged to Error_Log but never
    ''' surfaced in the dashboard's completion message. Per-item resolution failures inside
    ''' the loop had the same gap - logged, but never counted. Now returns a Tuple so the
    ''' caller can add this error count into its own running total instead of just checking
    ''' whether any rows came back.
    ''' </summary>
    Public Async Function RunAsync(days As Integer) As Task(Of (Rows As List(Of SphereMailStorageRow), ErrorCount As Integer))
        Dim results As New List(Of SphereMailStorageRow)
        Dim errorCount = 0

        Try
            Await SphereMailCustomersJob.RunAsync() ' original: Call Spheremail_Customers

            Dim token = Await SphereMailAuth.GetTokenAsync()
            Dim headers = New Dictionary(Of String, String) From {{"Authorization", token}}
            Dim queryParams = New Dictionary(Of String, String) From {{"limit", "1000"}}

            Dim response = Await ApiClient.GetAsync($"{AdminBaseUrl}/mail_items?shortcut=stored", queryParams, headers, timeoutSeconds:=15)
            If Not response.IsSuccess Then
                ErrorLogHelper.LogError("Spheremail Storage Emails Sub", "Error retrieving mail items from Spheremail")
                Return (results, 1)
            End If

            Dim data = response.DataAs(Of SphereMailMailItemsResponse)()

            For Each item In data.MailItems
                Dim createdAt = DateTime.Parse(item.ReceivedAt.Substring(0, 10)).Date

                If (DateTime.Now - createdAt).Days > days Then
                    Dim customerName = GetCustomerName(item.AccountNumber)

                    If String.IsNullOrEmpty(customerName) Then
                        ErrorLogHelper.LogError("Spheremail Storage Emails Sub", $"Unable to determine email address for PMB {item.AccountNumber}")
                        errorCount += 1
                        Continue For
                    End If

                    results.Add(New SphereMailStorageRow With {
                        .MailNumber = item.MailNumber,
                        .Location = Location,
                        .Customer = customerName,
                        .CreatedAt = createdAt,
                        .Sender = item.Sender,
                        .Quantity = item.Quantity,
                        .PrivateMailboxNumber = item.AccountNumber
                    })
                End If
            Next

        Catch ex As Exception
            ErrorLogHelper.LogError("Spheremail Storage Emails Sub", $"Unexpected error in Spheremail Storage Emails process: {ex.Message}")
            errorCount += 1
        End Try

        Return (results, errorCount)
    End Function

    ''' <summary>
    ''' Resolves a mailbox number to "CompanyName" via Customer_Ops -> Customer_QB
    ''' (the real SQL Server table name behind the "Customer_Sync_From_QB_SQL" alias -
    ''' see remarks above). Returns empty string if no match is found.
    ''' </summary>
    Private Function GetCustomerName(pmb As String) As String
        Const acctSql As String =
            "SELECT [Account Number] FROM Customer_Ops " &
            "WHERE [Company Mailbox] = @Pmb AND [Primary Office] = @Location"

        Dim acct As String = String.Empty
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(acctSql, conn)
                cmd.Parameters.AddWithValue("@Pmb", pmb)
                cmd.Parameters.AddWithValue("@Location", Location)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                acct = If(result Is Nothing OrElse result Is DBNull.Value, String.Empty, result.ToString())
            End Using
        End Using

        If String.IsNullOrEmpty(acct) Then Return String.Empty

        Const nameSql As String = "SELECT CompanyName FROM Customer_QB WHERE AccountNumber = @Acct"
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(nameSql, conn)
                cmd.Parameters.AddWithValue("@Acct", acct)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                Return If(result Is Nothing OrElse result Is DBNull.Value, String.Empty, result.ToString())
            End Using
        End Using
    End Function

End Module
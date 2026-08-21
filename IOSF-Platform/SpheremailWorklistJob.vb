Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command66_Click() ("Spheremail Worklist").
'''
''' REAL BUG FIXED: the original loops "For i = 1 To 7", but the Select Case only defines
''' Cases 1-6 with no Case Else. VBA leaves task/task2 unchanged when no case matches, so
''' iteration 7 silently re-runs the SAME API call as iteration 6 ("trash_requested") and
''' re-inserts duplicate "Trash" rows. Fixed here by looping 1 To 6 (matching the actual
''' number of defined cases) - not replicated, since there's no plausible reason to want
''' doubled trash entries on every run.
'''
''' Reuses SphereMailAuth.GetTokenAsync() for authentication, same as the other
''' api.spheremail.co-based jobs in this port.
'''
''' Forward-address lookup preserved exactly: only for "Forward" (i=1) and "Expedited
''' Forward" (i=5) items, a SECOND API call fetches the forward address's street. If that
''' second call fails, the original just leaves the address blank (no error, no abort) -
''' unlike the main mail_items call, whose failure IS a hard abort. This distinction is
''' preserved exactly, not homogenized.
'''
''' "Expd Frwd" task label gets the delivery-days count appended (e.g. "Expd Frwd 2 Day"),
''' matching the original's task2 = task2 & " " & Days & " Day" for i=5 only.
'''
''' received_at is truncated to its first 10 characters (an ISO-8601 date prefix from a
''' full timestamp string) before parsing, matching the original's own Left(received_at, 10).
'''
''' CustomerName ADDED per Al, after comparing directly against the original Access
''' report's own PDF output ("Spheremail Worklist") - the original groups/displays by
''' company name, not raw account number, which this port's earlier version omitted.
''' Looked up in one batched query against Customer_QB (by AccountNumber) after fetching
''' all rows, rather than per-row, to avoid one SQL round-trip per mail item.
'''
''' No temp table - Spheremail_Worklist_Temp was a local Access table (not a linked SQL
''' Server table, confirmed via the repo - no corresponding tbldefs entry, only a local
''' .sql/.xml pair), replaced here with an in-memory list - no database round-trip needed
''' for staging.
'''
''' PDF generation/printing now handled by ReportGenerator.GenerateSphereMailWorklistPdfAsync
''' (landscape, grouped by customer, alternating row shading - matching the original's own
''' report layout), opened directly with the system's default PDF viewer from
''' LandingPageForm - no intermediate grid window, per Al's explicit request. The earlier
''' SpheremailWorklistForm/DataTablePrinter-based approach is retired.
''' </summary>
Public Module SpheremailWorklistJob

    Private Const ApiBaseUrl As String = "https://api.spheremail.co/v1/admin"

    Private ReadOnly TaskShortcuts As String() = {
        "forward_requested", "envelope_picture_requested", "shred_requested",
        "scan_requested", "expedited_forward_requested", "trash_requested"
    }
    Private ReadOnly TaskLabels As String() = {
        "Forward", "Env Pic", "Shred", "Scan", "Expd Frwd", "Trash"
    }

    Public Async Function FetchWorklist() As Task(Of List(Of SpheremailWorklistRow))
        Dim rows As New List(Of SpheremailWorklistRow)

        Dim token = Await SphereMailAuth.GetTokenAsync()
        Dim headers = New Dictionary(Of String, String) From {{"Authorization", token}}

        For i = 0 To TaskShortcuts.Length - 1
            Dim shortcut = TaskShortcuts(i)
            Dim taskLabel = TaskLabels(i)
            Dim isForwardType = (i = 0 OrElse i = 4) ' Forward (i=0) or Expedited Forward (i=4), 0-indexed here vs the original's 1-indexed i=1/i=5

            Dim queryParams = New Dictionary(Of String, String) From {{"limit", "1000"}}
            Dim response = Await ApiClient.GetAsync($"{ApiBaseUrl}/mail_items?shortcut={shortcut}", queryParams, headers, timeoutSeconds:=15)
            If Not response.IsSuccess Then
                ' Matches the original's own hard abort (MsgBox + Exit Sub) on the main call failing.
                Throw New InvalidOperationException($"API Call Error fetching {shortcut} (status {CInt(response.StatusCode)}). Process Aborted.")
            End If

            Dim data = response.DataAs(Of SphereMailMailItemsResponse)()
            If data?.MailItems Is Nothing Then Continue For

            For Each item In data.MailItems
                Dim address = ""
                If isForwardType AndAlso Not String.IsNullOrEmpty(item.AccountId) AndAlso Not String.IsNullOrEmpty(item.ForwardAddressId) Then
                    ' Second call's failure is silently ignored, matching the original exactly - no error, no abort, just a blank address.
                    Try
                        Dim addrResponse = Await ApiClient.GetAsync($"{ApiBaseUrl}/customers/{item.AccountId}/forward_addresses/{item.ForwardAddressId}", Nothing, headers, timeoutSeconds:=15)
                        If addrResponse.IsSuccess Then
                            Dim addrData = addrResponse.DataAs(Of SphereMailForwardAddressResponse)()
                            address = If(addrData?.ForwardAddress?.Street, "")
                        End If
                    Catch
                        ' silently ignored, matching the original
                    End Try
                End If

                Dim displayTask = taskLabel
                If i = 4 Then displayTask = $"{taskLabel} {item.DeliveryDays} Day" ' Expedited Forward only

                Dim receivedDateText = If(item.ReceivedAt IsNot Nothing AndAlso item.ReceivedAt.Length >= 10, item.ReceivedAt.Substring(0, 10), item.ReceivedAt)
                Dim receivedDate As Date
                Date.TryParse(receivedDateText, receivedDate)

                rows.Add(New SpheremailWorklistRow With {
                    .MailNumber = item.MailNumber,
                    .AccountNumber = item.AccountNumber,
                    .ReceivedAt = receivedDate,
                    .Sender = item.Sender,
                    .Quantity = item.Quantity,
                    .Task = displayTask,
                    .Address = address
                })
            Next
        Next

        ApplyCustomerNames(rows)

        Return rows
    End Function

    ''' <summary>One batched lookup against Customer_QB for every distinct AccountNumber in the fetched rows, rather than a query per row.</summary>
    Private Sub ApplyCustomerNames(rows As List(Of SpheremailWorklistRow))
        Dim distinctAccounts = rows.Select(Function(r) r.AccountNumber).Where(Function(a) Not String.IsNullOrEmpty(a)).Distinct().ToList()
        If distinctAccounts.Count = 0 Then Return

        Dim names As New Dictionary(Of String, String)

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            conn.Open()
            For Each accountNumber In distinctAccounts
                Using cmd As New SqlCommand("SELECT CompanyName FROM Customer_QB WHERE AccountNumber = @AccountNumber", conn)
                    cmd.Parameters.AddWithValue("@AccountNumber", accountNumber)
                    Dim result = cmd.ExecuteScalar()
                    names(accountNumber) = If(result Is Nothing OrElse result Is DBNull.Value, "", result.ToString())
                End Using
            Next
        End Using

        For Each row In rows
            row.CustomerName = If(names.ContainsKey(row.AccountNumber), names(row.AccountNumber), "")
        Next
    End Sub

End Module
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
''' No temp table - Spheremail_Worklist_Temp was a local Access table (not a linked SQL
''' Server table, confirmed via the repo - no corresponding tbldefs entry, only a local
''' .sql/.xml pair), replaced here with an in-memory DataTable, same pattern already used
''' for other "_Temp" tables throughout this port (Customer_QB_Temp, Kube_Invoice_Temp,
''' etc.) - no database round-trip needed for staging.
'''
''' Printing (landscape, matching the original's explicit print orientation/duplex
''' settings) is handled by the reusable DataTablePrinter, from SpheremailWorklistForm's
''' own Print button - see that file.
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

    Public Async Function FetchWorklist() As Task(Of DataTable)
        Dim table As New DataTable()
        table.Columns.Add("mail_number")
        table.Columns.Add("account_number")
        table.Columns.Add("received_at", GetType(Date))
        table.Columns.Add("sender")
        table.Columns.Add("quantity")
        table.Columns.Add("task")
        table.Columns.Add("address")

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

                Dim row = table.NewRow()
                row("mail_number") = item.MailNumber
                row("account_number") = item.AccountNumber
                row("received_at") = receivedDate
                row("sender") = item.Sender
                row("quantity") = item.Quantity
                row("task") = displayTask
                row("address") = address
                table.Rows.Add(row)
            Next
        Next

        Return table
    End Function

End Module
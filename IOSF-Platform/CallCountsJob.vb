Imports Microsoft.Data.SqlClient

''' <summary>
''' Pulls reception-call records day by day from HostedSuite, aggregates them by
''' account/date/call-type, and syncs the totals into Call_Counts. Also flags customers
''' with no Evo_Customer_XRef mapping so they can be added.
'''
''' Date range is passed in explicitly by the caller - for an interactive run that means a
''' date picker; for a scheduled/batch run, GetNextStartDate() below computes a sensible
''' starting point automatically.
'''
''' The "missing Xref" check only warns about likely real customer accounts (ClientId
''' below 9000, parsed as an integer) rather than every unmapped ClientId, since internal/
''' high-numbered accounts aren't expected to have a mapping.
'''
''' The Authorization header for this API is computed at runtime from the "Evo Pass"
''' Config value (Base64("sanfran:" & EvoPass)) rather than stored as its own separate
''' Config entry - this keeps rotating the Evo password sufficient on its own, with no
''' second place that credential needs to be kept in sync.
''' </summary>
Public Module CallCountsJob

    Private Const ApiBaseUrl As String = "https://io2.hostedsuite.com/api/"
    Private Const PageSize As Integer = 1000

    Public Async Function RunAsync(startDate As Date, endDate As Date) As Task(Of Integer)
        Dim errorCount = 0
        Dim allCalls As New List(Of ReceptionCallItem)

        ' Fetch, day by day, paginated
        Dim currentDate = startDate
        While currentDate <= endDate
            Try
                allCalls.AddRange(Await FetchDayAsync(currentDate))
            Catch ex As Exception
                ErrorLogHelper.LogError("Call Counts to DB", $"Error retrieving calls for {currentDate:d}: {ex.Message}")
                errorCount += 1
            End Try
            currentDate = currentDate.AddDays(1)
        End While

        ' Load Evo_Customer_XRef into memory for the join (ClientId -> ThirdPartyAccountId)
        Dim xref = GetXrefLookup()

        ' Delete existing rows in this date range, then insert fresh aggregates - both in
        ' one transaction, so a mid-run failure can't leave the table with some accounts
        ' synced and others not for the same date range.
        Try
            ApplyCallCounts(startDate, endDate, allCalls, xref)
        Catch ex As Exception
            ErrorLogHelper.LogError("Call Counts to DB", $"SQL error applying call counts: {ex.Message}")
            errorCount += 1
        End Try

        ' Flag customers with no Xref mapping (real customers only - see remarks above)
        Dim missingXrefNames = allCalls.
            Where(Function(c) Not xref.ContainsKey(c.ClientId)).
            Where(Function(c) IsLikelyRealCustomer(c.ClientId)).
            Select(Function(c) c.ClientName).
            Distinct().
            ToList()

        For Each name In missingXrefNames
            ErrorLogHelper.LogError("Load Call Counts to DB", $"No Xref for Customer {name}")
            errorCount += 1
        Next

        Return errorCount
    End Function

    ''' <summary>Non-numeric ClientId values are treated as "not a real customer" (excluded) rather than causing a parse error.</summary>
    Private Function IsLikelyRealCustomer(clientId As String) As Boolean
        Dim numericId As Integer
        If Integer.TryParse(clientId, numericId) Then
            Return numericId < 9000
        End If
        Return False
    End Function

    Private Async Function FetchDayAsync(day As Date) As Task(Of List(Of ReceptionCallItem))
        Dim result As New List(Of ReceptionCallItem)
        Dim rangeStart = $"{day.Month}/{day.Day}/{day.Year} 00:00"
        Dim rangeEnd = $"{day.Month}/{day.Day}/{day.Year} 23:59"
        Dim billRange = $"{{Start:{rangeStart},End:{rangeEnd}}}"

        Dim headers = New Dictionary(Of String, String) From {{"Authorization", ComputeAuthHeader()}}
        Dim seenIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim page = 0
        Dim totalPages = 1

        While page < totalPages
            Dim queryParams = New Dictionary(Of String, String) From {
                {"DateCreated", billRange},
                {"StartTime", billRange},
                {"Page", page.ToString()},
                {"CountPerPage", PageSize.ToString()}
            }

            Dim response = Await ApiClient.GetAsync($"{ApiBaseUrl}reception-calls", queryParams, headers, timeoutSeconds:=20)
            If Not response.IsSuccess Then
                Throw New InvalidOperationException($"reception-calls page {page} for {day:d} returned {CInt(response.StatusCode)}")
            End If

            Dim data = response.DataAs(Of ReceptionCallsResponse)()
            totalPages = Math.Max(data.TotalPages, 1)

            If data.Items IsNot Nothing Then
                For Each item In data.Items
                    If seenIds.Add(item.Id) Then
                        item.ClientName = item.ClientName?.Replace("'", "''")
                        result.Add(item)
                    End If
                Next
            End If

            page += 1
        End While

        Return result
    End Function

    Private Function GetXrefLookup() As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)
        Const sql As String = "SELECT Id, ThirdPartyAccountId FROM Evo_Customer_XRef"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result(reader.GetString(0)) = reader.GetString(1)
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    ''' <summary>
    ''' Deletes the existing date range and inserts fresh aggregates in one transaction,
    ''' so a mid-run failure can't leave Call_Counts partially synced for that range.
    '''
    ''' Grouped by account number + call date + call type - NOT by client name, since more
    ''' than one ClientId (e.g. separate phone lines) can map to the same account number,
    ''' and grouping by name would split what should be one combined row into several,
    ''' each trying to insert the same primary key. Account number is parsed to Integer as
    ''' part of the grouping key itself (not just at insert time), so formatting
    ''' differences in the stored string (whitespace, leading zeros) can't cause the same
    ''' account to split into separate groups either. Call date is parsed to an actual
    ''' Date up front for the same reason - grouping by a raw string slice of the
    ''' timestamp is fragile if the API ever varies its string format.
    '''
    ''' IMPORTANT VB.NET DETAIL: the grouping key below is an anonymous type
    ''' ("New With {Key ..., Key ..., Key ...}"). The "Key" keyword on each field is
    ''' required for the type to get structural equality (two instances with the same
    ''' field values are treated as equal) - without it, VB.NET anonymous types fall back
    ''' to reference equality, and GroupBy would silently fail to combine any records at
    ''' all, no matter how obviously identical their account/date/type. Do not remove
    ''' "Key" from these fields.
    ''' </summary>
    Private Sub ApplyCallCounts(startDate As Date, endDate As Date, calls As List(Of ReceptionCallItem), xref As Dictionary(Of String, String))
        Dim matched = calls.Where(Function(c) xref.ContainsKey(c.ClientId))

        Dim grouped = matched.GroupBy(Function(c) New With {
            Key .AccountNum = CInt(xref(c.ClientId).Trim()),
            Key .CallDate = DateTime.Parse(c.StartTime).Date,
            Key .CallType = c.Type.Trim()
        }).ToList()

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            conn.Open()
            Using transaction = conn.BeginTransaction()
                Try
                    Using deleteCmd As New SqlCommand("DELETE FROM Call_Counts WHERE StartDate >= @StartDate AND EndDate <= @EndDate", conn, transaction)
                        deleteCmd.Parameters.AddWithValue("@StartDate", startDate)
                        deleteCmd.Parameters.AddWithValue("@EndDate", endDate)
                        deleteCmd.ExecuteNonQuery()
                    End Using

                    For Each g In grouped
                        Const insertSql As String =
                            "INSERT INTO Call_Counts (Account_Num, StartDate, EndDate, Company_Evo, Calls, Duration, Hold, Talk, Billable, CallType) " &
                            "VALUES (@AccountNum, @StartDate, @EndDate, @CompanyEvo, @Calls, @Duration, @Hold, @Talk, @Billable, @CallType)"

                        Using cmd As New SqlCommand(insertSql, conn, transaction)
                            cmd.Parameters.AddWithValue("@AccountNum", g.Key.AccountNum)
                            cmd.Parameters.AddWithValue("@StartDate", g.Key.CallDate)
                            cmd.Parameters.AddWithValue("@EndDate", g.Key.CallDate)
                            cmd.Parameters.AddWithValue("@CompanyEvo", g.First().ClientName) ' representative name for this combined account
                            cmd.Parameters.AddWithValue("@Calls", g.Count())
                            cmd.Parameters.AddWithValue("@Duration", g.Sum(Function(c) c.Duration) / 60)
                            cmd.Parameters.AddWithValue("@Hold", g.Sum(Function(c) c.HoldTime) / 60)
                            cmd.Parameters.AddWithValue("@Talk", g.Sum(Function(c) c.TalkTime) / 60)
                            cmd.Parameters.AddWithValue("@Billable", g.Sum(Function(c) c.TalkTime + c.TransferTime) / 60)
                            cmd.Parameters.AddWithValue("@CallType", g.Key.CallType)
                            cmd.ExecuteNonQuery()
                        End Using
                    Next

                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    ''' <summary>Computes the batch-mode start date: the day after the last recorded StartDate in Call_Counts, or 10 days ago if the table is empty.</summary>
    Public Function GetNextStartDate() As Date
        Const sql As String = "SELECT MAX(StartDate) FROM Call_Counts"
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return Date.Today.AddDays(-10)
                End If
                Return CDate(result).AddDays(1)
            End Using
        End Using
    End Function

End Module
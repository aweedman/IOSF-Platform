Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command18_Click() ("190.3 - Call Counts to DB").
''' Pulls reception-call records day by day from HostedSuite, aggregates them by
''' customer/date/call type, and syncs the totals into Call_Counts. Also flags customers
''' with no Evo_Customer_XRef mapping so they can be added.
'''
''' Changes from the VBA original:
'''  - Interactive date selection (the DatePicker/MsgBox confirmation loop) is NOT ported
'''    here - that's UI and belongs with the click handler. This function takes explicit
'''    startDate/endDate parameters; the caller computes them either interactively or,
'''    for batch runs, the same way the original did: startDate = (MAX(StartDate) in
'''    Call_Counts) + 1 day, endDate = yesterday.
'''  - Call_Counts_Temp was a local Access table (confirmed - no linked-table descriptor).
'''    Replaced by an in-memory List(Of ReceptionCallItem) - no DB round-trip, and no
'''    longer one DoCmd.RunSQL INSERT per individual call record (the original issued a
'''    separate INSERT for every single call - this collects them in memory and
'''    aggregates once).
'''  - THE "NO XREF" ERROR CHECK IS CHANGED, not just re-verified - see the conversation
'''    this was ported in. The original's WHERE clause
'''    (Evo_Customer_XRef_SQL.Id Is Null AND Evo_Customer_XRef_SQL.ThirdPartyAccountId < '9000')
'''    can never be true: ThirdPartyAccountId comes from the same LEFT-JOINed row as Id,
'''    so it's NULL whenever Id Is Null, and NULL < '9000' is never true in SQL's
'''    three-valued logic. Per your confirmed intent (only warn about missing Xrefs for
'''    real customers, not internal/high-numbered accounts), this now checks the RAW
'''    ClientId from the call record itself (before any Xref lookup) against 9000, since
'''    that's the only value available for a row that has no Xref match at all. Please
'''    sanity-check the 9000 threshold against real ClientId values once this runs -
'''    this is my best interpretation of the intent, not a verified fix.
'''  - Table names verified against tbldefs: Call_Counts_SQL -> real name Call_Counts,
'''    Evo_Customer_XRef_SQL -> real name Evo_Customer_XRef (both simple suffix-strips,
'''    confirmed via their linked-table .json descriptors).
'''  - CREDENTIAL FIX, then changed again per Al: the reception-calls API's Authorization
'''    header (a Basic-auth-style "IO base64string" credential) was originally hardcoded,
'''    then moved to a static "Call Counts Auth Header" Config value - but that value
'''    itself is a computed derivative of "Evo Pass" (confirmed by decoding the stored
'''    value: it's Base64("sanfran:" & EvoPassValue), and "BIg7%lY8" matched the real Evo
'''    Pass value exactly). Per Al, this is now computed fresh from "Evo Pass" at runtime
'''    instead of being stored as its own separate, pre-computed Config value - one fewer
'''    place a credential can go stale or leak independently, and rotating the Evo password
'''    alone is now sufficient (no separate manual re-encoding step into Config needed).
'''    The "sanfran" username is NOT itself a Config value - preserved as a literal
'''    constant, since Al only described the password portion as coming from Config.
'''  - Returns an error count instead of MsgBox/Batch-mode messaging - same pattern as
'''    every other job; the caller decides how to surface it.
''' </summary>
Public Module CallCountsJob

    Private Const ApiBaseUrl As String = "https://io2.hostedsuite.com/api/"
    Private Const PageSize As Integer = 1000 ' original: GC_Items

    Public Async Function RunAsync(startDate As Date, endDate As Date) As Task(Of Integer)
        Dim errorCount = 0
        Dim allCalls As New List(Of ReceptionCallItem)

        ' --- Fetch, day by day, paginated ---
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

        ' --- Load Evo_Customer_XRef into memory for the join (ClientId -> ThirdPartyAccountId) ---
        Dim xref = GetXrefLookup()

        ' --- Delete existing rows in this date range, then insert fresh aggregates -
        ' both in one transaction, so a mid-loop failure (like the duplicate-key one this
        ' was fixed after) can't leave the table in a partial state: some accounts synced,
        ' others not, for the same date range. ---
        Try
            ApplyCallCounts(startDate, endDate, allCalls, xref)
        Catch ex As Exception
            ErrorLogHelper.LogError("Call Counts to DB", $"SQL error applying call counts: {ex.Message}")
            errorCount += 1
        End Try

        ' --- Flag customers with no Xref mapping (real customers only - see remarks above) ---
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

    ''' <summary>
    ''' Interpretation of "ThirdPartyAccountId < 9000" applied to the raw ClientId instead
    ''' - see the class-level remarks. Non-numeric ClientId values are treated as "not a
    ''' real customer" (excluded) rather than causing a parse error.
    ''' </summary>
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
        ' Paging starts at Page=0 - reverted after a direct empirical test on
        ' VariableChargesToDbJob (same API family) confirmed the API is 0-indexed. An
        ' earlier version of this file switched to Page=1, based on a misdiagnosis: the
        ' duplicate-key error that prompted that change was actually caused by
        ' VariableChargesToDbJob's broken date-range filter (returning an unfiltered,
        ' overlapping set of charges), not by page indexing - this file's own date filter
        ' was already using the correct JSV format, so it was very likely working fine
        ' before that untested change. The Id-based dedup below is kept regardless, as a
        ' harmless safety net either way.
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
    ''' Deletes the existing range and inserts fresh aggregates in ONE transaction, so a
    ''' mid-loop failure can't leave Call_Counts with some accounts synced and others not
    ''' for the same date range - the original ran these as two independent statements
    ''' with no such guarantee, same class of gap already fixed elsewhere in this port
    ''' (CustomerMasterDeltaJob.ApplyDelta).
    '''
    ''' REAL BUG FIXED (not just re-verified) - see conversation this was fixed in.
    ''' Grouping by ClientName as well as AccountNum (matching the original's own
    ''' "GROUP BY ClientName, ThirdPartyAccountId, Date, CallType") is wrong: if two
    ''' different ClientIds (e.g. two phone lines/extensions) map to the SAME account
    ''' number in Evo_Customer_XRef but have different ClientName values, this produces
    ''' two separate groups that both try to INSERT the same (Account_Num, StartDate,
    ''' EndDate, CallType) primary key - a duplicate-key violation. This is inherited
    ''' directly from the original's own GROUP BY design, not something introduced by
    ''' porting - the original would have hit the identical crash for any account with
    ''' more than one registered line. Fixed here: group by account/date/call-type only,
    ''' and pick one representative name (the first one encountered) for the Company_Evo
    ''' display column on the combined row.
    '''
    ''' SECOND RELATED FIX, caught on the next test run against a different account:
    ''' grouping originally used the raw string ThirdPartyAccountId, converting to Integer
    ''' only at insert time - if that column has formatting inconsistencies across rows for
    ''' the same logical account (whitespace, leading zeros), two groups could still exist
    ''' that collapse to the same account number only once CInt() ran, causing the same
    ''' class of duplicate-key collision for a different account. Now converts to the
    ''' canonical Integer up front, as part of the grouping key itself.
    '''
    ''' THIRD RELATED FIX: ruled out duplicate Evo_Customer_XRef mappings directly (query
    ''' confirmed 0 accounts with more than one Id), so the remaining suspect was CallDate
    ''' itself - it was a raw string slice of StartTime (Substring(0,10)), grouped as a
    ''' STRING. If the reception-calls API ever returns StartTime in more than one format
    ''' (different separator, timezone suffix, whatever), two differently-formatted slices
    ''' could represent the same actual calendar date, group as DISTINCT keys in .NET, and
    ''' then both get inserted as @StartDate/@EndDate string parameters that SQL Server
    ''' parses down to the identical DATE value - a duplicate-key collision at the database
    ''' level despite .NET correctly seeing them as different groups. Parsing StartTime to
    ''' a real Date up front and grouping by THAT closes this regardless of source format
    ''' variability, and the SQL parameters are now proper typed Date values, not strings.
    ''' FIFTH AND ACTUAL ROOT-CAUSE FIX: the anonymous type used as the grouping key never
    ''' had the VB.NET "Key" keyword on its fields. In VB.NET (unlike C#), "New With {...}"
    ''' without "Key" on each field produces a MUTABLE anonymous type, and mutable
    ''' anonymous types do NOT get structural Equals/GetHashCode - they fall back to
    ''' reference equality. This meant GroupBy was NEVER actually aggregating anything:
    ''' every single call record became its own singleton group regardless of whether its
    ''' account/date/type matched another record's, confirmed directly via a pairwise
    ''' diagnostic showing grouped.Count matched the raw record count exactly, and every
    ''' pair of same-account calls flagged as "equal but kept separate". This explains
    ''' every symptom chased in this file's history: the primary-key violations (multiple
    ''' un-aggregated single-call "groups" sharing the same real key, each attempting its
    ''' own INSERT), and the earlier "DELETE removed 0 rows" result (no prior run had ever
    ''' successfully committed data into that range, since this same bug caused every
    ''' earlier attempt to hit an identical collision and roll back). None of the four
    ''' fixes above this comment were wrong, but none of them could have worked, since the
    ''' real defect was structural and sat underneath all of them.
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

    ''' <summary>
    ''' Computes the batch-mode start date: day after the last recorded StartDate in
    ''' Call_Counts. NOTE: the original's equivalent (DMax(...) + 1) would actually crash
    ''' with a type-mismatch error if Call_Counts has never had a row - Access's DMax
    ''' returns Null when the table's empty, and Null + 1 assigned to a Date variable
    ''' raises a runtime error in VBA. This defaults to 10 days ago instead (matching the
    ''' GC_Days = 10 convention already used elsewhere in Landing Page.cls) rather than
    ''' reproducing that crash.
    ''' </summary>
    Public Function GetNextStartDate() As Date
        Const sql As String = "SELECT MAX(StartDate) FROM Call_Counts"
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return Date.Today.AddDays(-10) ' table empty - original would have crashed here
                End If
                Return CDate(result).AddDays(1)
            End Using
        End Using
    End Function

End Module
Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command19_Click() ("150.1 - Variable Charges to DB").
'''
''' CUT OVER TO A NEW API per Al: the original called the OLDER io.hostedsuite.com
''' ListChargesRequest endpoint (POST, credentials sent as plain body parameters - no auth
''' header at all). This instead calls the newer io2.hostedsuite.com /api/charges endpoint
''' (GET, same Basic-auth-style Authorization header already used for Call Counts - see
''' HostedSuiteAuth), confirmed against that endpoint's own metadata
''' (io2.hostedsuite.com/api/json/metadata?op=ListCharges) before building this. Per Al,
''' more jobs are expected to cut over to this same new API family going forward.
'''
''' Field mapping from the new API's ChargeInfo to the old API's fields it replaces:
'''  - The new API splits the old "Service" field into ServiceId/ServiceName separately -
'''    ServiceName (the human-readable name) is what's stored in the Service column,
'''    matching what "Service" almost certainly meant in the old API. NOT independently
'''    confirmed against real data - worth checking on the first real run.
'''  - EntityStatus was read in the original but never actually used anywhere afterward
'''    (assigned to a variable, never referenced again) - not replicated, since it served
'''    no purpose in the original either.
'''
''' NOT INDEPENDENTLY VERIFIED: the exact query-string encoding for a DateRangeFilter
''' parameter (DateOfCharge.Start / DateOfCharge.End using ServiceStack's standard
''' dot-notation for flattened complex-type query params) and the date string format
''' expected (ISO 8601 assumed) - the metadata page documents the parameter shape but not
''' the wire format precisely. Worth confirming on the first real test run; if the API
''' rejects the date format or ignores the filter entirely, that's the first thing to
''' adjust.
'''
''' Date-range UI (DatePicker/MsgBox confirmation loop in the original) follows the exact
''' same pattern as Call Counts, per Al - see LandingPageForm's RunVariableCharges, which
''' uses the same DateRangeDialog + DefaultDateHelper.ComputeDefaultDate(26,-1)/(25,0)
''' defaults.
'''
''' Table name NOT independently verified: Variable_Charges_SQL -> assumed real name
''' Variable_Charges (the simple-strip convention that's held for most tables in this
''' port), but unlike Call_Counts/Evo_Customer_XRef this wasn't confirmed against a
''' tbldefs descriptor - worth checking if the DELETE/INSERT below fails with a
''' table-not-found error.
'''
''' Same "one transaction for the whole date range" robustness pattern as Call Counts
''' (see its remarks for the full reasoning): DELETE + all INSERTs run as one atomic unit,
''' so a mid-loop failure can't leave the range half-synced. This is a deliberate choice
''' to match Call Counts' established pattern (per Al's request that this job mirror it),
''' rather than the original's per-row "On Error Resume Next" partial-commit behavior.
''' </summary>
Public Module VariableChargesToDbJob

    Private Const ApiBaseUrl As String = "https://io2.hostedsuite.com/api/"
    Private Const PageSize As Integer = 1000

    Public Async Function RunAsync(startDate As Date, endDate As Date) As Task(Of Integer)
        Dim errorCount = 0
        Dim charges As New List(Of ChargeInfo)

        Try
            charges = Await FetchChargesAsync(startDate, endDate)
        Catch ex As Exception
            ErrorLogHelper.LogError("Variable Charges to DB", $"Error retrieving charges: {ex.Message}")
            Return 1
        End Try

        ' --- TEMPORARY DIAGNOSTIC - remove once confirmed the fetch is matching real data
        ' correctly. Distinguishes "API genuinely has 0 charges for this range" from
        ' "something is still silently wrong" - a run with no errors but no new rows
        ' either could mean either, and this makes it visible either way. ---
        ErrorLogHelper.LogError("Variable Charges DIAGNOSTIC",
            $"Fetched {charges.Count} charge(s) for range {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}.")
        ' --- END diagnostic ---

        Try
            ApplyVariableCharges(startDate, endDate, charges)
        Catch ex As Exception
            ErrorLogHelper.LogError("Variable Charges to DB", $"SQL error applying variable charges: {ex.Message}")
            errorCount += 1
        End Try

        Return errorCount
    End Function

    ''' <summary>
    ''' Paging starts at Page=0 - confirmed correct via a direct test (data came in with
    ''' Page=0 once the date-range encoding below was also fixed). An earlier version of
    ''' this function switched to Page=1, based on a misdiagnosis: the original
    ''' duplicate-key error was actually caused by the broken date-range encoding (which
    ''' returned an unfiltered, overlapping set of charges across many dates), not by
    ''' page indexing. The Id-based dedup below is kept regardless, as a harmless safety
    ''' net either way.
    ''' </summary>
    Private Async Function FetchChargesAsync(startDate As Date, endDate As Date) As Task(Of List(Of ChargeInfo))
        Dim result As New List(Of ChargeInfo)
        Dim seenIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Dim rangeStart = $"{startDate.Month}/{startDate.Day}/{startDate.Year} 00:00"
        Dim rangeEnd = $"{endDate.Month}/{endDate.Day}/{endDate.Year} 23:59"
        Dim dateRange = $"{{Start:{rangeStart},End:{rangeEnd}}}"

        Dim headers = New Dictionary(Of String, String) From {{"Authorization", HostedSuiteAuth.ComputeAuthHeader()}}
        Dim page = 0
        Dim totalPages = 1

        While page < totalPages
            Dim queryParams = New Dictionary(Of String, String) From {
                {"DateOfCharge", dateRange},
                {"Page", page.ToString()},
                {"CountPerPage", PageSize.ToString()}
            }

            Dim response = Await ApiClient.GetAsync($"{ApiBaseUrl}charges", queryParams, headers, timeoutSeconds:=30)
            response.EnsureSuccess()

            Dim data = response.DataAs(Of ChargesListResponse)()
            If data.Items IsNot Nothing Then
                For Each item In data.Items
                    If seenIds.Add(item.Id) Then result.Add(item)
                Next
            End If
            totalPages = Math.Max(data.TotalPages, 1)
            page += 1
        End While

        Return result
    End Function

    Private Sub ApplyVariableCharges(startDate As Date, endDate As Date, charges As List(Of ChargeInfo))
        ' --- TEMPORARY DIAGNOSTIC - remove once the duplicate-key mystery is resolved ---
        ' The in-memory dedup in FetchChargesAsync SHOULD make it impossible for `charges`
        ' to contain the same Id twice - if this fires, that dedup has a real bug. If it
        ' does NOT fire (charges is clean) but the PK violation still happens below, the
        ' problem is a pre-existing row in the table, not a duplicate within this fetch.
        Dim duplicateIdsWithinFetch = charges.GroupBy(Function(c) c.Id).Where(Function(g) g.Count() > 1).ToList()
        For Each dupe In duplicateIdsWithinFetch
            ErrorLogHelper.LogError("Variable Charges DIAGNOSTIC",
                $"Id '{dupe.Key}' appears {dupe.Count()} times in the fetched charges list itself - the in-memory dedup should have prevented this.")
        Next
        ' --- END diagnostic ---

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            conn.Open()
            Using transaction = conn.BeginTransaction()
                Try
                    Using deleteCmd As New SqlCommand("DELETE FROM Variable_Charges WHERE TransactionDate BETWEEN @StartDate AND @EndDate", conn, transaction)
                        deleteCmd.Parameters.AddWithValue("@StartDate", startDate)
                        deleteCmd.Parameters.AddWithValue("@EndDate", endDate)
                        deleteCmd.ExecuteNonQuery()
                    End Using

                    For Each charge In charges
                        Dim quantity = If(charge.Quantity.GetValueOrDefault() = 0, 1, charge.Quantity.GetValueOrDefault())
                        Dim transactionDate = DateTime.Parse(charge.DateOfCharge).Date

                        Const insertSql As String =
                            "INSERT INTO Variable_Charges (Id, ClientId, Company_Evo, Service, TransactionDate, Qty, Cost, Description) " &
                            "VALUES (@Id, @ClientId, @CompanyEvo, @Service, @TransactionDate, @Qty, @Cost, @Description)"

                        Try
                            Using cmd As New SqlCommand(insertSql, conn, transaction)
                                cmd.Parameters.AddWithValue("@Id", charge.Id)
                                cmd.Parameters.AddWithValue("@ClientId", charge.ClientId)
                                cmd.Parameters.AddWithValue("@CompanyEvo", If(charge.ClientName, String.Empty))
                                cmd.Parameters.AddWithValue("@Service", If(charge.ServiceName, String.Empty))
                                cmd.Parameters.AddWithValue("@TransactionDate", transactionDate)
                                cmd.Parameters.AddWithValue("@Qty", quantity)
                                cmd.Parameters.AddWithValue("@Cost", charge.Cost.GetValueOrDefault())
                                cmd.Parameters.AddWithValue("@Description", If(charge.Description, String.Empty))
                                cmd.ExecuteNonQuery()
                            End Using
                        Catch insertEx As Exception
                            ' --- TEMPORARY DIAGNOSTIC - remove once resolved ---
                            ' Checks, on the SAME connection/transaction, whether a row for
                            ' this exact Id already existed BEFORE this run - if so, its
                            ' stored TransactionDate tells us whether the DELETE's date
                            ' range is somehow missing it (e.g. a timezone/date-computation
                            ' mismatch between what THIS run computes and what was stored
                            ' previously).
                            Dim existingDetails As String = "(none found)"
                            Using checkCmd As New SqlCommand("SELECT TransactionDate FROM Variable_Charges WHERE Id = @Id", conn, transaction)
                                checkCmd.Parameters.AddWithValue("@Id", charge.Id)
                                Dim existingDate = checkCmd.ExecuteScalar()
                                If existingDate IsNot Nothing AndAlso existingDate IsNot DBNull.Value Then
                                    existingDetails = $"EXISTING row found with TransactionDate={CDate(existingDate):yyyy-MM-dd}"
                                End If
                            End Using

                            ErrorLogHelper.LogError("Variable Charges DIAGNOSTIC",
                                $"Insert failed for Id={charge.Id}, raw DateOfCharge='{charge.DateOfCharge}', computed TransactionDate={transactionDate:yyyy-MM-dd}, " &
                                $"ClientId={charge.ClientId}, ClientName={charge.ClientName}. This run's delete range: {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}. " &
                                $"{existingDetails}. Original error: {insertEx.Message}")
                            ' --- END diagnostic ---
                            Throw
                        End Try
                    Next

                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

End Module
Imports Microsoft.Data.SqlClient
Imports ClosedXML.Excel

''' <summary>
''' Direct port of Landing Page.cls: Command72_Click() ("190.1 - Kube Meetings to DB").
'''
''' NOT FIXED, FLAGGED FOR AL: RateCharged is used in the original's INSERT statement but
''' is NEVER assigned anywhere in the VBA - not read from Excel, not computed, nothing.
''' In VBA this evaluates as an empty string (an unassigned Variant defaults to Empty,
''' which concatenates as ""), meaning the original ALWAYS inserted an empty string for
''' RateCharged regardless of any actual rate. Replicated exactly here (a hardcoded empty
''' string constant) rather than guessed at, since it's not clear which Excel column (if
''' any) was originally meant to populate this.
'''
''' Duration/Length: the original computes Duration via Left(cell, 4) - a STRING
''' truncation to the first 4 characters, not numeric rounding - then embeds it UNQUOTED
''' in the INSERT (unlike every other text field, which is wrapped in quotes), strongly
''' implying Room_Usage.Length is a numeric column and the original author was treating
''' Duration as a number despite deriving it via string slicing. Replicated as: take the
''' first 4 characters as a string (matching Left(...,4) exactly), then parse that as a
''' decimal for the actual INSERT parameter.
'''
''' DayOfWeek: VBA's Weekday(date, vbMonday) returns 1=Monday...7=Sunday. .NET's
''' DateTime.DayOfWeek is 0=Sunday...6=Saturday - converted explicitly to match VBA's
''' numbering exactly, not .NET's own.
'''
''' ClientId lookup is a PREFIX match (Name LIKE 'ClientName%'), not an exact match -
''' preserved exactly as the original's own DLookup pattern, the only place in this port
''' Evo_Customer_XRef is queried this way (Kube's client names apparently carry extra
''' suffix text an exact match wouldn't catch). Matches "first row found" semantics (no
''' ORDER BY) if multiple rows share a prefix, same as the original DLookup's own
''' unordered behavior.
'''
''' DELETE-BEFORE-INSERT ADDED per Al (not in the original), same pattern as
''' SendProForwardsToDbJob: before processing any rows, deletes existing Room_Usage rows
''' whose StartTime falls within the MIN/MAX StartTime found among the rows that will
''' actually be inserted (i.e., AFTER the PaymentType/Cancelled filters below - a
''' filtered-out row was never going to be inserted, so its date shouldn't expand the
''' delete range). Makes re-running the same file (or an overlapping one) idempotent for
''' testing/reloading, instead of creating duplicates. This does NOT wrap the whole run in
''' one atomic transaction - the delete is its own quick upfront step, and the per-row
''' insert loop keeps its existing Resume-Next behavior (a failure on one row doesn't roll
''' back or stop the rest), matching the original's own per-row resilience design.
'''
''' Per-row failures are logged and do NOT stop the rest, matching the original's
''' On Error Resume Next - not the atomic all-or-nothing pattern used for Call
''' Counts/Variable Charges, since that wasn't requested for this job.
'''
''' Table name NOT independently verified: Room_Usage_SQL -> assumed real name
''' Room_Usage (the simple-strip convention that's held for most tables in this port, but
''' not confirmed against a tbldefs descriptor). Evo_Customer_XRef_SQL -> confirmed real
''' name Evo_Customer_XRef (already established elsewhere in this port).
'''
''' CenterName hardcoded to "San Francisco", matching the same "Burlingame decommissioned"
''' pattern already established in other SphereMail/RemoteLock jobs in this port - the
''' original had no location branching here at all, just a bare literal.
''' </summary>
Public Module KubeMeetingsToDbJob

    Private Const SheetName As String = "Report Data"
    Private Const FirstDataRow As Integer = 5
    Private Const GrandTotalMarker As String = "Grand Total:"

    Public Function RunAsync(excelFilePath As String) As Task(Of Integer)
        Return Task.Run(Function()
                            Dim errorCount = 0

                            ' Two-phase read: parse and filter every row up front, both to find the
                            ' min/max StartTime among rows that will actually be inserted (for the
                            ' upfront delete, per Al) and to stay consistent with the existing
                            ' "stop at Grand Total" loop condition, which needs to happen before either step.
                            Dim rows As List(Of KubeMeetingRow)
                            Try
                                Using workbook = New XLWorkbook(excelFilePath)
                                    Dim ws = workbook.Worksheet(SheetName)
                                    rows = ReadAndFilterRows(ws)
                                End Using
                            Catch ex As Exception
                                ' REAL GAP FIXED: this phase wasn't wrapped in a Try/Catch at all before -
                                ' an exception here would have propagated unhandled instead of being
                                ' logged, unlike every other job in this port, which always catches
                                ' fetch/read failures separately from SQL failures.
                                ErrorLogHelper.LogError("Kube Meetings to DB", $"Error reading Excel file: {ex.Message}")
                                Return 1
                            End Try

                            If rows.Count = 0 Then Return 0

                            Dim minDate = rows.Min(Function(r) r.StartTime)
                            Dim maxDate = rows.Max(Function(r) r.StartTime)

                            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                                conn.Open()

                                Try
                                    Using deleteCmd As New SqlCommand("DELETE FROM Room_Usage WHERE StartTime BETWEEN @MinDate AND @MaxDate", conn)
                                        deleteCmd.Parameters.AddWithValue("@MinDate", minDate)
                                        deleteCmd.Parameters.AddWithValue("@MaxDate", maxDate)
                                        deleteCmd.ExecuteNonQuery()
                                    End Using
                                Catch ex As Exception
                                    ErrorLogHelper.LogError("Kube Meetings to DB", $"SQL error in: DELETE FROM Room_Usage - {ex.Message}")
                                    errorCount += 1
                                End Try

                                For Each row In rows
                                    Try
                                        errorCount += ProcessRow(conn, row)
                                    Catch ex As Exception
                                        ErrorLogHelper.LogError("Kube Meetings to DB", $"SQL error: {ex.Message}")
                                        errorCount += 1
                                    End Try
                                Next
                            End Using

                            Return errorCount
                        End Function)
    End Function

    ''' <summary>Raw fields for one booking row, captured up front so the workbook only needs to be read once.</summary>
    Private Class KubeMeetingRow
        Public Property BookingId As String
        Public Property ClientName As String
        Public Property MeetingRoomName As String
        Public Property StartTime As Date
        Public Property EndTime As Date
        Public Property Duration As Decimal
    End Class

    ''' <summary>
    ''' REAL BUG FIXED: GetDateTime() requires a native Excel date value and throws
    ''' otherwise (confirmed via a real "Specified cast is not valid" error on the very
    ''' first data row). The original VBA used CDate(...), which is far more flexible -
    ''' it can also parse a text-formatted date string, not just a native date cell.
    ''' Tries the strict native read first (the common, fast case), falls back to parsing
    ''' the cell's string representation if that fails, matching CDate's own flexibility.
    ''' Confirmed necessary against the actual uploaded file - its Check-In/Check-Out
    ''' columns are text strings like "August 19, 2026 10:00 AM", not native date cells.
    ''' </summary>
    Private Function ReadDate(cell As IXLCell) As Date
        Try
            Return cell.GetDateTime()
        Catch
            Return Date.Parse(cell.GetString())
        End Try
    End Function

    ''' <summary>
    ''' Reads every row up front, stopping at the Grand Total marker, and applies the
    ''' PaymentType/Cancelled filters here (rather than per-row later) so the caller can
    ''' compute the delete-range date span from only the rows that will actually be
    ''' inserted.
    ''' </summary>
    Private Function ReadAndFilterRows(ws As IXLWorksheet) As List(Of KubeMeetingRow)
        Dim result As New List(Of KubeMeetingRow)
        Dim row = FirstDataRow

        While ws.Cell(row, 1).GetString() <> GrandTotalMarker
            Try
                Dim paymentType = ws.Cell(row, 14).GetString()
                Dim bookingStatus = ws.Cell(row, 13).GetString()

                ' REAL BUG FIXED (this time confirmed against the actual file, not
                ' guessed): CancelFee cells are currency-FORMATTED STRINGS like "$0.00",
                ' not native numeric values - TryGetValue(Of Decimal) still expects a
                ' numeric underlying type and fails against the dollar sign, same as the
                ' earlier GetValue(Of Decimal)() attempt did. Reading as a string and
                ' parsing with NumberStyles.Currency handles the "$" prefix correctly,
                ' matching VBA's own flexible implicit string-to-Currency coercion that
                ' let the original work fine against this same data.
                Dim cancelFeeText = ws.Cell(row, 15).GetString()
                Dim cancelFee As Decimal = 0
                If Not String.IsNullOrWhiteSpace(cancelFeeText) Then
                    Decimal.TryParse(cancelFeeText, Globalization.NumberStyles.Currency, Globalization.CultureInfo.InvariantCulture, cancelFee)
                End If

                If paymentType = "BillLater" AndAlso Not (bookingStatus = "Cancelled" AndAlso cancelFee = 0) Then
                    Dim durationRaw = ws.Cell(row, 11).GetString()
                    Dim durationText = durationRaw.Substring(0, Math.Min(4, durationRaw.Length)) ' Left(...,4) equivalent
                    Dim duration As Decimal
                    Decimal.TryParse(durationText, duration) ' original embedded this unquoted in SQL, implying a numeric column - defaults to 0 if the truncated text isn't a clean number

                    result.Add(New KubeMeetingRow With {
                        .BookingId = ws.Cell(row, 1).GetString(),
                        .ClientName = ws.Cell(row, 2).GetString().Replace("'", "''"),
                        .MeetingRoomName = ws.Cell(row, 6).GetString(),
                        .StartTime = ReadDate(ws.Cell(row, 9)),
                        .EndTime = ReadDate(ws.Cell(row, 10)),
                        .Duration = duration
                    })
                End If
            Catch ex As Exception
                ' Row number included so a remaining read failure (a different column
                ' than CancelFee) can be pinpointed exactly, rather than the generic
                ' "SQL error" the outer catch would otherwise show.
                Throw New InvalidOperationException($"Error reading Excel row {row}: {ex.Message}", ex)
            End Try

            row += 1
        End While

        Return result
    End Function

    ''' <summary>Returns 1 if this row was logged as an error (no matching ClientId), 0 otherwise.</summary>
    Private Function ProcessRow(conn As SqlConnection, row As KubeMeetingRow) As Integer
        Dim dayOfWeek = (CInt(row.StartTime.DayOfWeek) + 6) Mod 7 + 1 ' VBA vbMonday numbering - see class remarks

        Dim clientId As String = Nothing
        Using cmd As New SqlCommand("SELECT TOP 1 Id FROM Evo_Customer_XRef WHERE Name LIKE @ClientName + '%'", conn)
            cmd.Parameters.AddWithValue("@ClientName", row.ClientName)
            Dim result = cmd.ExecuteScalar()
            If result IsNot Nothing AndAlso result IsNot DBNull.Value Then clientId = result.ToString()
        End Using

        If String.IsNullOrEmpty(clientId) Then
            ErrorLogHelper.LogError("Kube Meetings to DB", $"Client ID not found for  {row.ClientName}")
            Return 1
        End If

        ' NOT FIXED - see class remarks. Always empty, matching the original exactly.
        Const rateCharged As String = ""

        Const insertSql As String =
            "INSERT INTO Room_Usage (Id, MeetingRoomName, CenterName, StartTime, EndTime, Length, ClientId, Company_Evo, DayOfWeek, RateCharged) " &
            "VALUES (@Id, @MeetingRoomName, 'San Francisco', @StartTime, @EndTime, @Length, @ClientId, @CompanyEvo, @DayOfWeek, @RateCharged)"

        Using cmd As New SqlCommand(insertSql, conn)
            cmd.Parameters.AddWithValue("@Id", row.BookingId)
            cmd.Parameters.AddWithValue("@MeetingRoomName", row.MeetingRoomName)
            cmd.Parameters.AddWithValue("@StartTime", row.StartTime)
            cmd.Parameters.AddWithValue("@EndTime", row.EndTime)
            cmd.Parameters.AddWithValue("@Length", row.Duration)
            cmd.Parameters.AddWithValue("@ClientId", clientId)
            cmd.Parameters.AddWithValue("@CompanyEvo", row.ClientName)
            cmd.Parameters.AddWithValue("@DayOfWeek", dayOfWeek)
            cmd.Parameters.AddWithValue("@RateCharged", rateCharged)
            cmd.ExecuteNonQuery()
        End Using

        Return 0
    End Function

End Module
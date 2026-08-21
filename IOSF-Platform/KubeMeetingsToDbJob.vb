Imports Microsoft.Data.SqlClient
Imports ClosedXML.Excel

''' <summary>
''' Imports Kube meeting-room bookings from an Excel report into the Room_Usage table.
'''
''' WORTH VERIFYING: RateCharged is always inserted as an empty string - no source column
''' for an actual rate has been identified yet. If a rate should be populated here, the
''' right Excel column for it needs to be confirmed first.
'''
''' Duration/Length: read as the first 4 characters of the cell's text (not numeric
''' rounding), then parsed as a decimal for the actual insert - Room_Usage.Length is a
''' numeric column, so this truncate-then-parse approach needs to keep producing a clean
''' number.
'''
''' DayOfWeek is stored using Monday=1...Sunday=7 numbering (not .NET's own
''' Sunday=0...Saturday=6), converted explicitly to match what's expected downstream.
'''
''' ClientId lookup is a PREFIX match (Name LIKE 'ClientName%'), not an exact match -
''' Kube's client names apparently carry extra suffix text that an exact match wouldn't
''' catch. If more than one row shares a prefix, the first one found (no particular order)
''' is used.
'''
''' Before processing any rows, existing Room_Usage rows are deleted for the StartTime
''' range found among the rows that will actually be inserted (after the PaymentType/
''' Cancelled filters below - a filtered-out row shouldn't expand the delete range),
''' making it safe to re-run the same file, or an overlapping one, without creating
''' duplicates. This isn't wrapped in one transaction with the row-by-row import - the
''' delete is its own quick upfront step, and each row's failure is handled independently
''' without rolling back the rest.
'''
''' CenterName is hardcoded to "San Francisco" - there's no other active location.
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
                            ' upfront delete) and to stay consistent with the "stop at Grand Total"
                            ' loop condition, which needs to happen before either step.
                            Dim rows As List(Of KubeMeetingRow)
                            Try
                                Using workbook = New XLWorkbook(excelFilePath)
                                    Dim ws = workbook.Worksheet(SheetName)
                                    rows = ReadAndFilterRows(ws)
                                End Using
                            Catch ex As Exception
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
    ''' Tries a native Excel date read first (the common, fast case), falling back to
    ''' parsing the cell's text representation if that fails - the Check-In/Check-Out
    ''' columns in this report can come through as text strings (e.g. "August 19, 2026
    ''' 10:00 AM") rather than native date cells.
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

                ' CancelFee cells are currency-formatted text (e.g. "$0.00"), not native
                ' numeric values - parsing with NumberStyles.Currency handles the "$"
                ' prefix correctly.
                Dim cancelFeeText = ws.Cell(row, 15).GetString()
                Dim cancelFee As Decimal = 0
                If Not String.IsNullOrWhiteSpace(cancelFeeText) Then
                    Decimal.TryParse(cancelFeeText, Globalization.NumberStyles.Currency, Globalization.CultureInfo.InvariantCulture, cancelFee)
                End If

                If paymentType = "BillLater" AndAlso Not (bookingStatus = "Cancelled" AndAlso cancelFee = 0) Then
                    Dim durationRaw = ws.Cell(row, 11).GetString()
                    Dim durationText = durationRaw.Substring(0, Math.Min(4, durationRaw.Length))
                    Dim duration As Decimal
                    Decimal.TryParse(durationText, duration) ' defaults to 0 if the truncated text isn't a clean number

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
                ' Row number included so a read failure can be pinpointed exactly, rather
                ' than a generic "SQL error" from the outer catch.
                Throw New InvalidOperationException($"Error reading Excel row {row}: {ex.Message}", ex)
            End Try

            row += 1
        End While

        Return result
    End Function

    ''' <summary>Returns 1 if this row was logged as an error (no matching ClientId), 0 otherwise.</summary>
    Private Function ProcessRow(conn As SqlConnection, row As KubeMeetingRow) As Integer
        Dim dayOfWeek = (CInt(row.StartTime.DayOfWeek) + 6) Mod 7 + 1 ' Monday=1...Sunday=7 - see class remarks

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

        Const rateCharged As String = "" ' see class remarks

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
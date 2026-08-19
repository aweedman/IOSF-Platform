Imports Microsoft.Data.SqlClient
Imports ClosedXML.Excel
Imports System.Windows.Forms

''' <summary>
''' Direct port of Landing Page.cls: Command22_Click() ("Room Usage Report").
'''
''' Rather than replicate the original's deeply nested Access SQL (three levels of
''' subqueries with SWITCH/IIF-computed columns) as one large T-SQL statement, this
''' fetches the joined raw data with a simpler query and computes every derived field
''' (Day name, RateCharged fallback, AfterHours, Low/High clamped times, DayRate,
''' AdditionalHours) in VB.NET instead - easier to verify against real output, and
''' consistent with how the rest of this port handles complex derived logic (e.g.
''' CallCountsJob's aggregation).
'''
''' Confirmed with Al before building:
'''  - Customer_Header_SQL is the same table as Customer_Ops_Header (already used
'''    elsewhere in this port).
'''  - The account-number filter (< 9000) uses a NUMERIC comparison here, not the
'''    original's string comparison ("< '9000'") - Al confirmed account numbers are
'''    always 4 characters and numeric is fine, and this is also consistent with how
'''    CustomerMasterForm's own gallery filter already works.
'''  - The AdditionalHours calculation's trailing Replace(value, "0", "", 1, 1) - Access's
'''    string-replace removing the FIRST "0" character anywhere in the number's string
'''    form - is NOT replicated. As written it would incorrectly mangle a genuine value
'''    like "10" into "1", and Al couldn't recall its original intent without seeing
'''    unmodified output first. AdditionalHours here is a plain decimal, no string
'''    manipulation. Revisit if Al finds the numbers need a specific format after seeing
'''    this version's output.
'''
''' RateCharged fallback logic (LIKE pattern matching on MeetingRoomName when
''' Room_Usage.RateCharged isn't exactly 1 character) means the RateCharged-never-
''' assigned bug flagged in KubeMeetingsToDbJob is effectively harmless for THIS report -
''' an empty RateCharged (length 0) always falls through to the room-name-based fallback
''' anyway.
'''
''' AfterHours/DayRate/Low/High logic, and the DayOfWeek->name mapping (1=Monday...
''' 7=Sunday), preserved exactly per the original's own SWITCH conditions, evaluated in
''' the same order (first match wins, matching Access's SWITCH semantics).
'''
''' The LEFT JOIN to Customer_QB combined with filtering on its AccountNumber in the
''' WHERE clause effectively makes this an INNER join in practice (a Room_Usage row with
''' no matching Customer_QB account gets excluded) - this matches the original's own
''' behavor exactly, not a deviation, since Access SQL behaves the same way.
'''
''' Output: an .xlsx file (via ClosedXML) to the user's Desktop, matching the original's
''' TransferSpreadsheet export location and intent (a folder named "Room_Events" in the
''' original - this creates "Room_Events.xlsx" as a single file there instead, since a
''' query export naturally produces one workbook, not a folder).
''' </summary>
Public Module RoomUsageReportJob

    Public Function RunAsync(fromDate As Date, toDate As Date) As Task(Of Integer)
        Return Task.Run(Function()
                            Try
                                Dim rows = FetchRows(fromDate, toDate)
                                Dim outputPath = WriteReport(rows)
                                ' Shown directly from the job, same precedent as SendProForwardsToDbJob's
                                ' own summary popup - this report's signature matches every other job's
                                ' plain Task(Of Integer) return (needed for LandingPageForm's standard
                                ' RunJobAsync helper), so the output file path is surfaced here rather
                                ' than via a non-standard return value.
                                MessageBox.Show($"Report complete. Saved to:{Environment.NewLine}{outputPath}", "Room Usage Report")
                                Return 0
                            Catch ex As Exception
                                ErrorLogHelper.LogError("Room Usage Report", $"Error generating report: {ex.Message}")
                                Return 1
                            End Try
                        End Function)
    End Function

    Private Class ReportRow
        Public Property FullName As String
        Public Property CenterName As String
        Public Property MeetingRoomName As String
        Public Property StartTime As DateTime
        Public Property EndTime As DateTime
        Public Property Member As String
        Public Property RawRateCharged As String
        Public Property Length As Decimal
        Public Property DayOfWeek As Integer ' 1=Monday...7=Sunday, matching KubeMeetingsToDbJob's own storage convention
        Public Property IsHoliday As Boolean
    End Class

    Private Function FetchRows(fromDate As Date, toDate As Date) As List(Of ReportRow)
        Dim result As New List(Of ReportRow)

        Const sql As String =
            "SELECT Q.FullName, RU.CenterName, RU.MeetingRoomName, RU.StartTime, RU.EndTime, " &
            "COH.Member, RU.RateCharged, RU.Length, RU.DayOfWeek, " &
            "CASE WHEN H.[Date] IS NOT NULL THEN 1 ELSE 0 END AS IsHoliday " &
            "FROM Room_Usage RU " &
            "LEFT JOIN Evo_Customer_XRef X ON RU.ClientId = X.Id " &
            "LEFT JOIN Customer_Ops_Header COH ON X.ThirdPartyAccountId = CAST(COH.Account_Num AS VARCHAR(20)) " &
            "LEFT JOIN Customer_QB Q ON X.ThirdPartyAccountId = Q.AccountNumber " &
            "LEFT JOIN Holidays H ON CAST(RU.StartTime AS DATE) = H.[Date] " &
            "WHERE TRY_CAST(Q.AccountNumber AS INT) < 9000 " &
            "AND CAST(RU.StartTime AS DATE) >= @FromDate AND CAST(RU.StartTime AS DATE) <= @ToDate " &
            "ORDER BY Q.FullName, CAST(RU.StartTime AS DATE)"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@FromDate", fromDate.Date)
                cmd.Parameters.AddWithValue("@ToDate", toDate.Date)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result.Add(New ReportRow With {
                            .FullName = If(reader.IsDBNull(0), "", reader.GetString(0)),
                            .CenterName = If(reader.IsDBNull(1), "", reader.GetString(1)),
                            .MeetingRoomName = If(reader.IsDBNull(2), "", reader.GetString(2)),
                            .StartTime = reader.GetDateTime(3),
                            .EndTime = reader.GetDateTime(4),
                            .Member = If(reader.IsDBNull(5), "", reader.GetString(5)),
                            .RawRateCharged = If(reader.IsDBNull(6), "", reader.GetString(6)),
                            .Length = If(reader.IsDBNull(7), 0D, Convert.ToDecimal(reader.GetValue(7))),
                            .DayOfWeek = Convert.ToInt32(reader.GetValue(8)),
                            .IsHoliday = reader.GetInt32(9) = 1
                        })
                    End While
                End Using
            End Using
        End Using

        Return result
    End Function

    Private ReadOnly DayNames As String() = {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}
    Private ReadOnly BusinessStart As New TimeSpan(8, 30, 0)
    Private ReadOnly BusinessEnd As New TimeSpan(17, 0, 0)

    Private Function WriteReport(rows As List(Of ReportRow)) As String
        Using workbook = New XLWorkbook()
            Dim ws = workbook.Worksheets.Add("Room Events")

            Dim headers = {"FullName", "CenterName", "MeetingRoomName", "StartTime", "EndTime", "Day", "ReservationDate", "Member", "RateCharged", "Length", "DayRate", "AdditionalHours"}
            For i = 0 To headers.Length - 1
                ws.Cell(1, i + 1).Value = headers(i)
            Next

            Dim outRow = 2
            For Each row In rows
                Dim dayName = DayNames(row.DayOfWeek - 1)

                ' RateCharged: use Room_Usage's own value only if it's exactly 1
                ' character, otherwise fall back to a room-name pattern match. Since
                ' RateCharged is always blank in practice (see class remarks), this
                ' effectively always falls through to the pattern match for Kube-sourced
                ' rows.
                Dim rateCharged As String
                If row.RawRateCharged IsNot Nothing AndAlso row.RawRateCharged.Length = 1 Then
                    rateCharged = row.RawRateCharged
                ElseIf row.MeetingRoomName.Contains("Office") Then
                    rateCharged = "O"
                ElseIf row.MeetingRoomName.StartsWith("Conf") Then
                    rateCharged = "C"
                ElseIf row.MeetingRoomName.StartsWith("Meeting") Then
                    rateCharged = "M"
                ElseIf row.MeetingRoomName.StartsWith("Workstation") Then
                    rateCharged = "W"
                Else
                    rateCharged = ""
                End If

                Dim startTod = row.StartTime.TimeOfDay
                Dim endTod = row.EndTime.TimeOfDay

                ' 'D' for holiday/weekend (Saturday=6, Sunday=7 in this DayOfWeek
                ' numbering) takes precedence over 'T' (partial after-hours), matching
                ' Access's SWITCH() evaluating conditions in order and returning the
                ' first match.
                Dim afterHours As String = Nothing
                If row.IsHoliday OrElse row.DayOfWeek = 6 OrElse row.DayOfWeek = 7 Then
                    afterHours = "D"
                ElseIf startTod < BusinessStart OrElse endTod > BusinessEnd Then
                    afterHours = "T"
                End If

                Dim low = If(startTod < BusinessStart, BusinessStart, startTod)
                Dim high = If(endTod > BusinessEnd, BusinessEnd, endTod)
                Dim inHoursSpanHours = (high - low).TotalHours

                Dim dayRate As String = Nothing
                If (afterHours Is Nothing AndAlso row.Length >= 6) OrElse (afterHours = "T" AndAlso inHoursSpanHours >= 6) Then
                    dayRate = "X"
                End If

                Dim additionalHours As Decimal? = Nothing
                If dayRate = "X" Then
                    Dim afterHigh = (endTod - high).TotalHours ' always >= 0, since High = Min(BusinessEnd, EndTime)
                    Dim beforeLow = (low - startTod).TotalHours ' always >= 0, since Low = Max(BusinessStart, StartTime)
                    Dim inHoursExtra = If(inHoursSpanHours > 8, inHoursSpanHours - 8, 0)
                    additionalHours = CDec(afterHigh + beforeLow + inHoursExtra)
                End If

                ws.Cell(outRow, 1).Value = row.FullName
                ws.Cell(outRow, 2).Value = row.CenterName
                ws.Cell(outRow, 3).Value = row.MeetingRoomName
                ws.Cell(outRow, 4).Value = row.StartTime.ToString("h:mm tt")
                ws.Cell(outRow, 5).Value = row.EndTime.ToString("h:mm tt")
                ws.Cell(outRow, 6).Value = dayName
                ws.Cell(outRow, 7).Value = row.StartTime.Date
                ws.Cell(outRow, 7).Style.DateFormat.Format = "M/d/yyyy"
                ws.Cell(outRow, 8).Value = row.Member
                ws.Cell(outRow, 9).Value = rateCharged
                ws.Cell(outRow, 10).Value = row.Length
                ws.Cell(outRow, 11).Value = dayRate
                If additionalHours.HasValue Then ws.Cell(outRow, 12).Value = additionalHours.Value

                outRow += 1
            Next

            ws.Columns().AdjustToContents()

            Dim outputPath = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Room_Events.xlsx")
            workbook.SaveAs(outputPath)
            Return outputPath
        End Using
    End Function

End Module
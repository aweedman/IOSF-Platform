Imports Microsoft.Data.SqlClient
Imports System.Text

''' <summary>
''' Direct port of Functions.bas: AfterHours().
''' Finds RemoteLock entries outside business hours, excluding employees and known
''' non-badge codes, and emails a summary for anything that looks like a customer
''' accessing a room after hours.
'''
''' Changes from the VBA original (behavior-preserving unless noted):
'''  - Timezone conversion now uses TimeZoneHelper (local, no API call) instead of
'''    api.ipgeolocation.io - see TimeZoneHelper.vb. The artificial 10-second wait
'''    loop that throttled the old API call is gone; nothing replaced it because
'''    nothing calls out to a rate-limited service anymore.
'''  - employees/holidays lookups (UBound(Filter(...))) became HashSet(Of String)
'''    and List(Of DateTime) with .Contains - same semantics, cleaner code.
'''  - The Facilities Code -> customer name DLookup is now a parameterized query
'''    (was string-concatenated in the original). Also added ISNULL() around each
'''    half of the concatenation: Access's "&" silently treats NULL as empty string,
'''    but SQL Server's "+" returns NULL for the whole expression if either side is
'''    NULL. Without ISNULL(), a customer with a blank CompanyName would produce a
'''    blank Name line in .NET where the VBA version showed just the contact name -
'''    this fix restores the original's actual behavior on SQL Server.
'''  - Batch/InputBox interactive prompt became a daysBack parameter - wire this to
'''    a command-line arg for interactive runs, or default to 1 for Task Scheduler.
''' </summary>
Public Module AfterHoursJob

    Private ReadOnly RoomNames As New Dictionary(Of String, String) From {
        {"93d7ac51-fa84-4b92-a26b-c0e90c0005e5", "Day Office 1"},
        {"5058ae19-7de7-4c0a-82e2-389123779dfa", "Day Office 2"},
        {"bb60581e-f976-4eda-b6af-b1896e5070d6", "Meeting Room"},
        {"6693c7aa-42a4-44cc-b517-67a1c972fe9f", "Conference Room"}
    }

    Private Class RemoteLockEvent
        Public Property OccurredAt As DateTime
        Public Property Pin As String
        Public Property PublisherId As String
    End Class

    Public Async Function RunAsync(Optional daysBack As Integer = 1) As Task
        Try
            Dim holidays = GetHolidayDates()
            Dim employeePins = GetEmployeePins()
            Dim minDate = DateTime.Today.AddDays(-daysBack - 1)
            Dim events = GetRemoteLockEvents(minDate)

            Dim message As New StringBuilder()

            For Each evt In events
                ' stop entirely once events are further back than DaysBack + 2 (raw occurred_at, pre-conversion)
                If (DateTime.Today - evt.OccurredAt.Date).Days > daysBack + 2 Then
                    Exit For
                End If

                ' skip blank/day-code pins
                If evt.Pin = "00000" OrElse evt.Pin = "12345" Then Continue For

                ' skip employees
                If employeePins.Contains(evt.Pin) Then Continue For

                Dim converted = TimeZoneHelper.ConvertUtcToPacific(evt.OccurredAt)
                Dim convertedDate = converted.Date

                ' skip entries that land on today once converted to Pacific
                If convertedDate = DateTime.Today Then Continue For

                ' stop entirely once converted date is further back than DaysBack
                If (DateTime.Today - convertedDate).Days > daysBack Then
                    Exit For
                End If

                ' skip weekday, non-holiday entries that fall inside business hours (8:30 AM - 5:00 PM)
                Dim isWeekday = converted.DayOfWeek <> DayOfWeek.Saturday AndAlso converted.DayOfWeek <> DayOfWeek.Sunday
                Dim isHoliday = holidays.Contains(convertedDate)
                Dim withinBusinessHours = converted.TimeOfDay > TimeSpan.FromHours(8.5) AndAlso converted.TimeOfDay < TimeSpan.FromHours(17)
                If isWeekday AndAlso Not isHoliday AndAlso withinBusinessHours Then Continue For

                Dim roomName As String = If(RoomNames.ContainsKey(evt.PublisherId), RoomNames(evt.PublisherId), "Unknown")
                Dim customerName = GetCustomerName(evt.Pin)

                If message.Length = 0 Then
                    message.AppendLine("The following are yesterday's after-hours room accesses. Please check cameras and enter in Evo as needed.")
                    message.AppendLine()
                End If

                message.AppendLine($"Room: {roomName}")
                message.AppendLine($"Name: {customerName}")
                message.AppendLine($"Entry Time: {converted:M/d/yyyy h:mm tt}")
                message.AppendLine()
            Next

            If message.Length > 0 Then
                Dim toUser = ConfigHelper.GetConfigValue("Email Afterhours Meetings User")
                EmailHelper.SendEmail(toUser, "After Hours Meeting Room Entries", message.ToString())
            End If

        Catch ex As Exception
            EmailHelper.EmailError("Undefined error in RemoteLock after hours entries process")
        End Try
    End Function

    Private Function GetHolidayDates() As HashSet(Of DateTime)
        Dim result As New HashSet(Of DateTime)
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand("SELECT Date FROM Holidays", conn)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result.Add(reader.GetDateTime(0).Date)
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Function GetEmployeePins() As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand("SELECT user_name FROM IO_Employees", conn)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result.Add(reader.GetString(0))
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Function GetRemoteLockEvents(minDate As DateTime) As List(Of RemoteLockEvent)
        Dim result As New List(Of RemoteLockEvent)
        Const sql As String =
            "SELECT occurred_at, pin, publisher_id FROM RemoteLock_Events " &
            "WHERE occurred_at >= @MinDate ORDER BY occurred_at DESC"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@MinDate", minDate)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result.Add(New RemoteLockEvent With {
                            .OccurredAt = reader.GetDateTime(0),
                            .Pin = reader.GetString(1),
                            .PublisherId = reader.GetString(2)
                        })
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Function GetCustomerName(pin As String) As String
        Const sql As String =
            "SELECT ISNULL([Contact Name], '') + ', ' + ISNULL(CompanyName, '') AS FullName " &
            "FROM Customer_Ops_All WHERE [Facilities Code] = @Pin"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Pin", pin)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                Return If(result Is Nothing OrElse result Is DBNull.Value, String.Empty, result.ToString())
            End Using
        End Using
    End Function

End Module
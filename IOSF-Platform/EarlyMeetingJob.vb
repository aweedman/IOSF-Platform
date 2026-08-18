''' <summary>
''' Replaces the original HostedSuite-based EarlyMeeting() with a check against public
''' Google Calendar ICS feeds - one per meeting room - since IOSF no longer uses
''' HostedSuite's own calendar/reservation system; all rooms now sync to separate Google
''' Calendars (confirmed with Al).
'''
''' Design decisions (confirmed with Al):
'''  - Reads each room's PUBLIC ICS feed URL directly (no OAuth, no API key needed at
'''    all). Each calendar's sharing must be set to "public" in Google Calendar's own
'''    settings for this to work - a calendar that's ever made private again will simply
'''    fail to fetch (404/403), logged as an error for that one room like any other
'''    per-room failure, without stopping the rest.
'''  - The room list lives in Config as "Early Meeting Calendar - {RoomName}" rows (see
'''    ConfigHelper.GetConfigValuesByPrefix) rather than being hardcoded here, so rooms
'''    can be added or removed later by editing Config alone - no code change needed.
'''  - "Same logic" as the original (confirmed with Al): a SINGLE combined earliest
'''    meeting across ALL rooms in the 8:30-9:00 AM window, one email - not a separate
'''    check per room.
'''  - The original's "Misc." room-name exclusion has NO direct equivalent here - none of
'''    the new rooms are literally named "Misc.", and nothing in this request implied
'''    filtering by event title, so no such filter is applied. If some placeholder/
'''    non-meeting event type on these calendars turns out to need excluding by title,
'''    that needs its own explicit rule from Al - not invented here without evidence.
'''  - Event times are converted to Pacific via the same TimeZoneHelper.ConvertUtcToPacific
'''    already used by AfterHoursJob, for consistency with the rest of the app.
'''
''' REQUIRES ADDING THE Ical.Net NUGET PACKAGE to the project (Visual Studio's NuGet
''' Package Manager, or `dotnet add package Ical.Net`) - used because it correctly expands
''' recurring events (RRULE) within a date range; a hand-rolled ICS parser would very
''' likely mishandle this, and conference-room calendars commonly have recurring bookings.
'''
''' NOTE: this file could not be compile-tested against the real Ical.Net library in the
''' sandbox this was written in (NuGet isn't reachable from there). The first version's
''' GetOccurrences(DateTime, DateTime) call didn't match the installed version's API
''' (confirmed via a real build error: the resolved overload expected (CalDateTime,
''' EvaluationOptions), not two DateTimes) - fixed by using the single-argument
''' GetOccurrences(CalDateTime) form instead, with the end of the window enforced manually
''' via Exit For rather than guessing at EvaluationOptions' shape too. Still not
''' compile-verified end to end - if GetOccurrences(CalDateTime) itself doesn't match
''' either, that's the next place to check.
''' </summary>
Public Module EarlyMeetingJob

    Private Const ConfigPrefix As String = "Early Meeting Calendar - "

    Public Async Function RunAsync() As Task
        Try
            Dim calendars = ConfigHelper.GetConfigValuesByPrefix(ConfigPrefix)
            If calendars.Count = 0 Then
                ErrorLogHelper.LogError("Early Meeting", $"No room calendars configured (expected Config rows named '{ConfigPrefix}{{Room}}')")
                Return
            End If

            ' Always checks TOMORROW's window, not today's - confirmed with Al this job is
            ' meant to give advance notice of the next day's first meeting (e.g. run the
            ' evening before), not report on a window that may have already passed today.
            Dim targetDate = DateTime.Today.AddDays(1)
            Dim startWindow = targetDate.AddHours(8).AddMinutes(30) ' 8:30 AM Pacific, tomorrow
            Dim endWindow = targetDate.AddHours(9)                  ' 9:00 AM Pacific, tomorrow

            Dim earliestTime As DateTime? = Nothing

            For Each kvp In calendars
                Dim roomName = kvp.Key
                Dim icsUrl = kvp.Value

                Try
                    Dim response = Await ApiClient.GetAsync(icsUrl, timeoutSeconds:=20)
                    response.EnsureSuccess()

                    Dim calendar = Ical.Net.Calendar.Load(response.Body)
                    ' REAL BUG FIXED: DateTime.Today has Kind=Local (and everything derived
                    ' from it via AddHours/AddMinutes preserves that Kind) - but Ical.Net's
                    ' CalDateTime constructor only accepts Kind=Utc or Kind=Unspecified,
                    ' confirmed via a real runtime error ("An instance of CalDateTime can
                    ' only be initialized from a DateTime of kind Utc or Unspecified").
                    ' Fixed by re-tagging just the Kind (not the actual date/time value) to
                    ' Unspecified right before construction - this doesn't affect the
                    ' startPacific/startWindow/endWindow comparisons elsewhere in this
                    ' function, since DateTime's comparison operators compare the raw ticks
                    ' value only and ignore Kind entirely.
                    Dim startCal As New Ical.Net.DataTypes.CalDateTime(DateTime.SpecifyKind(startWindow, DateTimeKind.Unspecified))
                    Dim occurrences = calendar.GetOccurrences(startCal)

                    For Each occurrence In occurrences
                        Dim startUtc = occurrence.Period.StartTime.AsUtc
                        Dim startPacific = TimeZoneHelper.ConvertUtcToPacific(startUtc)

                        ' Occurrences come out in chronological order - once we're past the
                        ' window, stop entirely rather than continuing to iterate (this
                        ' matters for infinitely-recurring events, which would otherwise
                        ' enumerate forever).
                        If startPacific >= endWindow Then Exit For

                        If startPacific >= startWindow Then
                            If earliestTime Is Nothing OrElse startPacific < earliestTime.Value Then
                                earliestTime = startPacific
                            End If
                        End If
                    Next

                Catch ex As Exception
                    ' One room's calendar failing to fetch/parse shouldn't stop checking
                    ' the others - log it and continue, same log-and-continue philosophy
                    ' already used elsewhere in this port for per-item failures.
                    ErrorLogHelper.LogError("Early Meeting", $"Error reading calendar for room '{roomName}': {ex.Message}")
                End Try
            Next

            If earliestTime Is Nothing Then Return

            Dim subject = $"First Meeting at {earliestTime.Value:h:mm tt}"
            Dim toUser = ConfigHelper.GetConfigValue("Email Meetings User")
            EmailHelper.SendEmail(toUser, subject, String.Empty)

        Catch ex As Exception
            ' Matches the original's overall behavior: an unexpected failure (e.g. Config
            ' unreachable) gets logged/emailed rather than propagated - confirmed with Al
            ' this stays as-is for this job (see conversation re: AfterHours' identical
            ' pattern).
            EmailHelper.EmailError("Error in Early Meeting notification process")
        End Try
    End Function

End Module
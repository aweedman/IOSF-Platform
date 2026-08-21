''' <summary>
''' Computes a sensible default date for pre-filling an interactive date-picker dialog -
''' e.g. "the 26th of last month" as a billing-cycle start date default.
'''
''' dayOfMonth=31 is a sentinel meaning "the last day of the month" rather than a literal
''' 31st (so it works correctly for months with fewer days). monthOffset shifts from
''' today's month (negative for past months, positive for future months, 0 for the
''' current month) before applying dayOfMonth.
''' </summary>
Public Module DefaultDateHelper

    Public Function ComputeDefaultDate(dayOfMonth As Integer, monthOffset As Integer) As Date
        Dim dateCalc = Date.Today

        If monthOffset <> 0 Then
            dateCalc = dateCalc.AddMonths(monthOffset)
        End If

        If dayOfMonth = 31 Then
            ' "Last day of the month" sentinel: advance a month, snap to the 1st, then
            ' step back one day - always lands on the correct last day regardless of
            ' month length.
            dateCalc = dateCalc.AddMonths(1)
            dateCalc = New Date(dateCalc.Year, dateCalc.Month, 1)
            dateCalc = dateCalc.AddDays(-1)
        Else
            Dim dayOffset = dayOfMonth - dateCalc.Day
            dateCalc = dateCalc.AddDays(dayOffset)
        End If

        Return dateCalc
    End Function

End Module
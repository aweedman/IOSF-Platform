''' <summary>
''' Replicates Landing Page.cls's local "DatePicker" Sub's default-date computation - NOT
''' the modDatePicker/InputDateField module (that's the different, superseded-by-
''' DateTimePicker one). This one is a small utility used to compute a SENSIBLE DEFAULT
''' date to pre-fill an interactive date dialog with - e.g. "26th of last month" for a
''' billing-cycle start date. The original:
'''
'''   Public Sub DatePicker(ReturnDate As Variant, ByVal Day As Integer, ByVal MonthOffset
'''   As Integer, ByVal Header As String)
'''
''' computes a default via (today, shifted by MonthOffset months, then normalized to the
''' given Day-of-month; Day=31 is a sentinel meaning "last day of the month" instead of a
''' literal 31st), then opens a picker pre-filled with that default. This port splits that
''' into two pieces, matching the separation-of-concerns used throughout this port: this
''' function just computes the default DATE VALUE; the actual picker UI is a plain
''' DateRangeDialog (or similar) that the caller pre-fills with it.
'''
''' Known call sites in the original (add more here as other interactive dialogs get
''' built):
'''   Command18_Click (Call Counts): DatePicker(BillStartDate, 26, -1, ...) / (BillEndDate, 25, 0, ...)
'''     -> 26th of last month through 25th of this month (a fixed billing-cycle default)
'''   Spheremail_Storage's interactive Days=999 branch: DatePicker(BillStartDate, 1, 1, "Invoice Date")
'''     -> 1st of NEXT month (not yet wired into any dashboard button - flagging for when it is)
''' </summary>
Public Module DefaultDateHelper

    Public Function ComputeDefaultDate(dayOfMonth As Integer, monthOffset As Integer) As Date
        Dim dateCalc = Date.Today

        If monthOffset <> 0 Then
            dateCalc = dateCalc.AddMonths(monthOffset)
        End If

        If dayOfMonth = 31 Then
            ' Sentinel meaning "last day of the month" rather than a literal 31st -
            ' matches the original's "add a month, set to the 1st, subtract a day" trick.
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
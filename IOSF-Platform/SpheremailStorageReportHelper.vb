''' <summary>
''' Direct port of Landing Page.cls: Command47_Click() ("Spheremail Storage Report"), plus
''' the Days=999 interactive branch of the shared Spheremail_Storage(Days) subroutine that
''' this button calls with Days=999.
'''
''' Reuses SphereMailStorageJob.RunAsync(days) directly - the SAME underlying logic
''' already built for Spheremail Storage Emails, confirmed by Al's own recollection and
''' the actual source (both call the same shared VBA subroutine). This file only adds the
''' piece that subroutine's own port deliberately left out: computing "Days" from an
''' interactively-picked Invoice Date, rather than always taking a fixed Days value.
'''
''' Days computation preserved exactly: BillStartDate (the picked Invoice Date) minus 3
''' months, minus 1 day, snapped to the 26th of that resulting month, then Days =
''' DateDiff("d", that date, Now).
'''
''' Default date for the picker (1st of NEXT month) uses DefaultDateHelper.
''' ComputeDefaultDate(1, 1) - a positive month offset, unlike every other call site in
''' this port so far, but already supported by that helper (its own doc comment already
''' flagged this exact call site before this job was wired up).
''' </summary>
Public Module SpheremailStorageReportHelper

    Public Function ComputeDaysFromInvoiceDate(invoiceDate As Date) As Integer
        Dim billStartDate = invoiceDate.AddMonths(-3).AddDays(-1)
        billStartDate = New Date(billStartDate.Year, billStartDate.Month, 26)
        Return CInt((DateTime.Now - billStartDate).TotalDays)
    End Function

    Public Function RowsToDataTable(rows As List(Of SphereMailStorageRow)) As DataTable
        Dim table As New DataTable()
        table.Columns.Add("MailNumber")
        table.Columns.Add("Location")
        table.Columns.Add("Customer")
        table.Columns.Add("CreatedAt", GetType(Date))
        table.Columns.Add("Sender")
        table.Columns.Add("Quantity")
        table.Columns.Add("PrivateMailboxNumber")

        For Each r In rows
            Dim row = table.NewRow()
            row("MailNumber") = r.MailNumber
            row("Location") = r.Location
            row("Customer") = r.Customer
            row("CreatedAt") = r.CreatedAt
            row("Sender") = r.Sender
            row("Quantity") = r.Quantity
            row("PrivateMailboxNumber") = r.PrivateMailboxNumber
            table.Rows.Add(row)
        Next

        Return table
    End Function

End Module
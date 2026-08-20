''' <summary>
''' Direct port of Landing Page.cls: Command47_Click() ("Spheremail Storage Report"), plus
''' the Days=999 interactive branch of the shared Spheremail_Storage(Days) subroutine that
''' this button calls with Days=999.
'''
''' REVISED per Al: originally displayed results in a custom DataGridView with its own
''' print support (DataTablePrinter), but that print logic used equal-width columns with
''' no text clipping, causing long values to visually overlap into adjacent columns
''' (confirmed via a real printed PDF Al sent). Rather than fix that print logic, this now
''' reuses ReportGenerator.GenerateSphereMailStoragePdfAsync directly - the SAME,
''' already-proven PDF generator SphereMailStorageEmailJob uses for its email
''' attachments (proportional column widths, correct sorting/date formatting, matches the
''' original Access report's actual layout). The generated PDF is opened directly with the
''' system's default viewer, matching the original's own DoCmd.OpenReport ...,
''' acViewPreview - no intermediate grid step, per Al's explicit request.
'''
''' Days computation preserved exactly: BillStartDate (the picked Invoice Date) minus 3
''' months, minus 1 day, snapped to the 26th of that resulting month, then Days =
''' DateDiff("d", that date, Now).
''' </summary>
Public Module SpheremailStorageReportHelper

    Public Function ComputeDaysFromInvoiceDate(invoiceDate As Date) As Integer
        Dim billStartDate = invoiceDate.AddMonths(-3).AddDays(-1)
        billStartDate = New Date(billStartDate.Year, billStartDate.Month, 26)
        Return CInt((DateTime.Now - billStartDate).TotalDays)
    End Function

End Module
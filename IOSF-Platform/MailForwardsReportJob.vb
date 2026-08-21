Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms

''' <summary>
''' Direct port of Landing Page.cls: Command62_Click() ("Mail Forwards Report").
'''
''' Markup formula uses the SHARED MailForwardMarkup constants (MarkupPercentage,
''' MarkupCap, RoundingIncrement), per Al - the same ones MailForwardsToEvoJob uses, so
''' the two can't drift apart and either can be updated in one place. See
''' MailForwardMarkup.vb.
'''
''' ONLY ONE DATE IS PROMPTED (unlike most other billing-cycle jobs): the end date is
''' COMPUTED from the start date, matching the original's own DateSerial(Year(DateAdd("m",
''' 1, FromDate)), Month(DateAdd("m", 1, FromDate)), 25) - the 25th of the month AFTER
''' FromDate's month. Not a separate prompt.
'''
''' FedEx branch uses FedEx.Total_Cost for the Postage/Markup calculation, matching
''' MailForwardsToEvoJob's own FedEx branch exactly - confirmed with Al after finding a
''' real discrepancy (this report originally used SendPro.Total_Cost here instead, which
''' would have disagreed with the actual posted Evo charge for the same shipment). Fixed
''' to match, not left as a silent inconsistency between the two jobs.
'''
''' Carrier display label preserved exactly (a DELIBERATE difference from
''' MailForwardsToEvoJob's own filtering logic, confirmed correct by Al - RY6026 is
''' genuinely a UPS account): a shipment with Carrier_Acct = 'RY6026' is included in the
''' same USPS-like filter/markup group as actual USPS, but is DISPLAYED here as "UPS"
''' (IIF(Carrier_Acct = 'RY6026', 'UPS', 'USPS')), not "USPS".
'''
''' UNION (not UNION ALL) preserved exactly, unlike MailForwardsToEvoJob's own UNION ALL -
''' this report deduplicates identical rows across its two branches; that job does not.
''' Both preserved as genuinely different, not homogenized.
'''
''' Same "no file export, opens directly in a grid" pattern as CopierCountsReportJob/
''' CallCountsReportJob - a read-only grid dialog, not an Excel export.
'''
''' Table names: SendPro, FedEx, Customer_QB all confirmed real tables elsewhere in this
''' port already.
''' </summary>
Public Module MailForwardsReportJob

    Public Function FetchReport(fromDate As Date) As DataTable
        Dim toDate = New Date(fromDate.AddMonths(1).Year, fromDate.AddMonths(1).Month, 25)

        Const sql As String = "
            SELECT CompanyName, Account_Num, Class, Carrier, Tracking_Num, Transaction_Date, Company_Sender, Recipient, Company_Recipient,
                Postage, Markup, ROUND((Postage + Markup) / @RoundIncrement, 0) * @RoundIncrement AS Amount_Charged_Rounded
            FROM (
                SELECT Customer_QB.CompanyName, SendPro.Account_Num, SendPro.Class,
                    IIF(SendPro.Carrier_Acct = 'RY6026', 'UPS', 'USPS') AS Carrier,
                    SendPro.Tracking_Num, SendPro.Transaction_Date, SendPro.Company_Sender,
                    SendPro.Recipient, SendPro.Company AS Company_Recipient, SendPro.Total_Cost AS Postage,
                    IIF(SendPro.Total_Cost * @MarkupPct < @MarkupCap, ROUND(SendPro.Total_Cost * @MarkupPct, 2), @MarkupCap) AS Markup
                FROM Customer_QB
                INNER JOIN SendPro ON Customer_QB.AccountNumber = CAST(SendPro.Account_Num AS VARCHAR(20))
                WHERE Customer_QB.IsActive = 1 AND SendPro.Transaction_Date >= @FromDate AND SendPro.Transaction_Date <= @ToDate
                AND SendPro.SM_Status <> 'Voided' AND SendPro.SM_Status NOT LIKE '%Refund%'
                AND (UPPER(SendPro.Carrier) = 'USPS' OR SendPro.Carrier_Acct = 'RY6026')

                UNION

                SELECT Customer_QB.CompanyName, SendPro.Account_Num, SendPro.Class, 'FedEx' AS Carrier,
                    SendPro.Tracking_Num, SendPro.Transaction_Date, SendPro.Company_Sender,
                    SendPro.Recipient, SendPro.Company AS Company_Recipient, FedEx.Total_Cost AS Postage,
                    IIF(FedEx.Total_Cost * @MarkupPct < @MarkupCap, ROUND(FedEx.Total_Cost * @MarkupPct, 2), @MarkupCap) AS Markup
                FROM Customer_QB
                INNER JOIN SendPro ON Customer_QB.AccountNumber = CAST(SendPro.Account_Num AS VARCHAR(20))
                INNER JOIN FedEx ON SendPro.Tracking_Num = FedEx.Tracking_Num
                WHERE Customer_QB.IsActive = 1 AND FedEx.Billing_Start_Date = @FromDate
            ) AS Union_Table
            ORDER BY CompanyName, Transaction_Date"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using adapter As New SqlDataAdapter(sql, conn)
                adapter.SelectCommand.Parameters.AddWithValue("@FromDate", fromDate)
                adapter.SelectCommand.Parameters.AddWithValue("@ToDate", toDate)
                adapter.SelectCommand.Parameters.AddWithValue("@MarkupPct", MailForwardMarkup.MarkupPercentage)
                adapter.SelectCommand.Parameters.AddWithValue("@MarkupCap", MailForwardMarkup.MarkupCap)
                adapter.SelectCommand.Parameters.AddWithValue("@RoundIncrement", MailForwardMarkup.RoundingIncrement)
                Dim table As New DataTable()
                adapter.Fill(table)
                Return table
            End Using
        End Using
    End Function

End Module

''' <summary>Simple read-only grid dialog for displaying MailForwardsReportJob's results.</summary>
Public Class MailForwardsReportForm
    Inherits Form

    Public Sub New(fromDate As Date)
        Text = "Mail Forwards Report"
        Width = 1100
        Height = 600
        StartPosition = FormStartPosition.CenterScreen

        Dim grid As New DataGridView With {
            .Dock = DockStyle.Fill,
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            .SelectionMode = DataGridViewSelectionMode.CellSelect
        }
        AddHandler grid.DataError, Sub(sender, e) e.ThrowException = False

        Controls.Add(grid)

        Try
            Dim table = MailForwardsReportJob.FetchReport(fromDate)
            grid.DataSource = table

            ' Per Al: the three amount columns (last three in the select list) show only
            ' two decimal places, rather than whatever raw precision the underlying
            ' money/decimal columns carry.
            For Each colName In {"Postage", "Markup", "Amount_Charged_Rounded"}
                If grid.Columns.Contains(colName) Then
                    grid.Columns(colName).DefaultCellStyle.Format = "F2"
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(Me, $"Error running report: {ex.Message}", "Mail Forwards Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms

''' <summary>
''' Direct port of Landing Page.cls: Command59_Click() ("190.4 - Call Counts Report").
''' Same pattern as CopierCountsReportJob (same file for both the query and its display
''' form) - the original opens results directly in Access's own datasheet grid rather
''' than exporting to a file, so this is a read-only grid dialog rather than a file
''' export.
'''
''' Aggregate query preserved exactly: per (Account_Num, CompanyName, Service_Level)
''' group, sums Calls/Billable/Duration within the selected date range, ordered by
''' CompanyName (this report has an explicit ORDER BY; Copier Counts Report did not -
''' preserved as an actual difference between the two, not homogenized).
'''
''' Join casting preserved exactly: the Customer_QB join casts Account_Num to a string to
''' match Customer_QB.AccountNumber (a string column, confirmed elsewhere in this port),
''' while the Customer_Ops_Header join does NOT cast - both Call_Counts.Account_Num and
''' Customer_Ops_Header.Account_Num appear to be numeric already (Customer_Ops_Header's
''' Account_Num is confirmed integer elsewhere in this port), matching the original's own
''' lack of a CStr() there.
'''
''' Table names: Call_Counts (from Call_Counts_SQL) was already confirmed as a real SQL
''' Server table earlier in this port. Customer_Ops_Header (from Customer_Header_SQL) and
''' Customer_QB (from Customer_Sync_From_QB_SQL) are both confirmed real tables elsewhere
''' in this port already.
''' </summary>
Public Module CallCountsReportJob

    Public Function FetchReport(fromDate As Date, toDate As Date) As DataTable
        Const sql As String = "
            SELECT Call_Counts.Account_Num AS AccountNum, Customer_QB.CompanyName AS CompanyName, Customer_Ops_Header.Service_Level AS ServiceLevel,
                SUM(Call_Counts.Calls) AS Calls, SUM(Call_Counts.Billable) AS Billable, SUM(Call_Counts.Duration) AS Duration
            FROM Call_Counts
            INNER JOIN Customer_QB ON CAST(Call_Counts.Account_Num AS VARCHAR(20)) = Customer_QB.AccountNumber
            LEFT JOIN Customer_Ops_Header ON Call_Counts.Account_Num = Customer_Ops_Header.Account_Num
            WHERE Call_Counts.StartDate >= @FromDate AND Call_Counts.StartDate <= @ToDate
            GROUP BY Call_Counts.Account_Num, Customer_QB.CompanyName, Customer_Ops_Header.Service_Level
            ORDER BY Customer_QB.CompanyName"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using adapter As New SqlDataAdapter(sql, conn)
                adapter.SelectCommand.Parameters.AddWithValue("@FromDate", fromDate)
                adapter.SelectCommand.Parameters.AddWithValue("@ToDate", toDate)
                Dim table As New DataTable()
                adapter.Fill(table)
                Return table
            End Using
        End Using
    End Function

End Module

''' <summary>Simple read-only grid dialog for displaying CallCountsReportJob's results.</summary>
Public Class CallCountsReportForm
    Inherits Form

    Public Sub New(fromDate As Date, toDate As Date)
        Text = "Call Counts Report"
        Width = 900
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
            grid.DataSource = CallCountsReportJob.FetchReport(fromDate, toDate)
        Catch ex As Exception
            MessageBox.Show(Me, $"Error running report: {ex.Message}", "Call Counts Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
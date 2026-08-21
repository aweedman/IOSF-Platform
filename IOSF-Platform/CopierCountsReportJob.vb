Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms

''' <summary>
''' Shows the Copier Counts report in a read-only grid dialog: per (Customer Name, user
''' full_name, location) group, sums Print/Copy pages split into Black & White
''' (total_pages - total_color_pages) and Color (total_color_pages) buckets, plus a
''' separate Scan pages total - all within the selected date range.
''' </summary>
Public Module CopierCountsReportJob

    Public Function FetchReport(fromDate As Date, toDate As Date) As DataTable
        Const sql As String = "
            SELECT DISTINCT Customer_QB.Name, Printer_Usage_Log.full_name, Printer_Usage_Log.location,
                SUM(IIF(Printer_Usage_Log.job_type = 'PRINT' OR Printer_Usage_Log.job_type = 'COPY', Printer_Usage_Log.total_pages - Printer_Usage_Log.total_color_pages, 0)) AS [Print BW],
                SUM(IIF(Printer_Usage_Log.job_type = 'PRINT' OR Printer_Usage_Log.job_type = 'COPY', Printer_Usage_Log.total_color_pages, 0)) AS [Print Color],
                SUM(IIF(Printer_Usage_Log.job_type = 'SCAN', Printer_Usage_Log.total_pages, 0)) AS Scan
            FROM Printer_Usage_Log
            INNER JOIN Customer_Ops_Item ON Printer_Usage_Log.user_name = Customer_Ops_Item.Fac_Code
            INNER JOIN Customer_QB ON CAST(Customer_Ops_Item.Account_Num AS VARCHAR(20)) = Customer_QB.AccountNumber
            WHERE Printer_Usage_Log.usage_day BETWEEN @FromDate AND @ToDate
            GROUP BY Customer_QB.Name, Printer_Usage_Log.full_name, Printer_Usage_Log.location"

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

''' <summary>Simple read-only grid dialog for displaying CopierCountsReportJob's results.</summary>
Public Class CopierCountsReportForm
    Inherits Form

    Public Sub New(fromDate As Date, toDate As Date)
        Text = "Copier Counts Report"
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
            grid.DataSource = CopierCountsReportJob.FetchReport(fromDate, toDate)
        Catch ex As Exception
            MessageBox.Show(Me, $"Error running report: {ex.Message}", "Copier Counts Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
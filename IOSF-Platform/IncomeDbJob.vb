Imports Microsoft.Data.SqlClient
Imports System.Data.Odbc

''' <summary>
''' Direct port of Landing Page.cls: Command48_Click() ("Income to DB").
''' Re-pulls QuickBooks' IncomeByCustomerDetail report for a date range and replaces
''' matching rows in SQL Server - a full delete-and-reinsert refresh, not a delta.
'''
''' Changes from the VBA original:
'''  - The original ran through a two-step Access-specific trick: a QueryDef named
'''    "SP_Shell" had its .SQL property rewritten at runtime to the sp_report statement,
'''    then queried as "SELECT ... FROM SP_Shell". That indirection existed only because
'''    Access needed a query OBJECT to select from - it's not meaningful outside Access.
'''    This just executes the sp_report statement directly via OdbcCommand and reads the
'''    result set - no equivalent indirection needed.
'''  - IMPORTANT: SP_Shell's own export showed a DIFFERENT QODBC DSN ("QuickBooks Data 64")
'''    than the one used elsewhere in this port ("QuickBooks Data 64-Bit QRemote", used by
'''    GetNextInvoiceNumber/QbCustomerRepository). Confirmed this is the same underlying
'''    QuickBooks connection, just named differently at different points in time - both
'''    now use ConfigHelper.QodbcConnectionString, no separate property needed.
'''  - Interactive date selection (DatePicker/MsgBox loop) is NOT ported here - that's UI.
'''    This function takes explicit fromDate/toDate; the caller computes them, matching
'''    the original's batch-mode default (first day of current month through last day).
'''  - The original's DELETE ran under "On Error Resume Next" with no logging at all if it
'''    failed - not reproduced, consistent with the same fix already applied in
'''    CustomerMasterDeltaJob. Delete failures are now logged like everything else.
'''  - Table name verified: IncomeByCustomerDetail_SQL -> real SQL Server name
'''    IncomeByCustomerDetail_QB (NOT a simple suffix-strip - confirmed via its tbldefs
'''    .json SourceTableName, same pattern as the earlier Customer_Sync_From_QB_SQL case).
'''  - Returns an error count instead of MsgBox/Batch-mode messaging - same pattern as
'''    every other job; the caller decides how to surface it.
''' </summary>
Public Module IncomeDbJob

    Public Async Function RunAsync(fromDate As Date, toDate As Date) As Task(Of Integer)
        Dim errorCount = 0

        Try
            DeleteExistingRange(fromDate, toDate)
        Catch ex As Exception
            ErrorLogHelper.LogError("Upload Income", $"Error deleting existing range: {ex.Message}")
            errorCount += 1
        End Try

        Dim rows As List(Of IncomeReportRow)
        Try
            rows = Await FetchReportAsync(fromDate, toDate)
        Catch ex As Exception
            ErrorLogHelper.LogError("Upload Income", $"SQL error in sp_report call: {ex.Message}")
            Return errorCount + 1
        End Try

        Try
            InsertRows(rows)
        Catch ex As Exception
            ErrorLogHelper.LogError("Upload Income", $"SQL error inserting income rows: {ex.Message}")
            errorCount += 1
        End Try

        Return errorCount
    End Function

    Private Function FetchReportAsync(fromDate As Date, toDate As Date) As Task(Of List(Of IncomeReportRow))
        Return Task.Run(Function()
            Dim result As New List(Of IncomeReportRow)
            Dim fromDateFormat = fromDate.ToString("yyyy-MM-dd")
            Dim toDateFormat = toDate.ToString("yyyy-MM-dd")

            Dim sql =
                "sp_report IncomeByCustomerDetail show TxnType, RefNumber, Date, NameAccountNumber, Account, Class, Amount " &
                $"parameters ReportBasis = 'Cash', AccountFilterType = 'OrdinaryIncome', DateFrom = {{d'{fromDateFormat}'}}, " &
                $"DateTo = {{d'{toDateFormat}'}} where Text is Null and Blank is Null"

            Using conn As New OdbcConnection(ConfigHelper.QodbcConnectionString)
                Using cmd As New OdbcCommand(sql, conn)
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            result.Add(New IncomeReportRow With {
                                .TxnType = GetString(reader, "TxnType"),
                                .RefNumber = GetString(reader, "RefNumber"),
                                .TxnDate = GetDateTime(reader, "Date"),
                                .NameAccountNumber = GetString(reader, "NameAccountNumber"),
                                .Account = GetString(reader, "Account"),
                                .TxnClass = GetString(reader, "Class"),
                                .Amount = GetDecimal(reader, "Amount")
                            })
                        End While
                    End Using
                End Using
            End Using

            Return result
        End Function)
    End Function

    Private Sub DeleteExistingRange(fromDate As Date, toDate As Date)
        Const sql As String = "DELETE FROM IncomeByCustomerDetail_QB WHERE Date BETWEEN @FromDate AND @ToDate"
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@FromDate", fromDate)
                cmd.Parameters.AddWithValue("@ToDate", toDate)
                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub InsertRows(rows As List(Of IncomeReportRow))
        Const sql As String =
            "INSERT INTO IncomeByCustomerDetail_QB (TxnType, RefNumber, Date, NameAccountNumber, Account, Class, Amount) " &
            "VALUES (@TxnType, @RefNumber, @Date, @NameAccountNumber, @Account, @Class, @Amount)"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            conn.Open()
            For Each row In rows
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@TxnType", CType(If(row.TxnType, DBNull.Value), Object))
                    cmd.Parameters.AddWithValue("@RefNumber", CType(If(row.RefNumber, DBNull.Value), Object))
                    cmd.Parameters.AddWithValue("@Date", CType(If(row.TxnDate.HasValue, row.TxnDate.Value, DBNull.Value), Object))
                    cmd.Parameters.AddWithValue("@NameAccountNumber", CType(If(row.NameAccountNumber, DBNull.Value), Object))
                    cmd.Parameters.AddWithValue("@Account", CType(If(row.Account, DBNull.Value), Object))
                    cmd.Parameters.AddWithValue("@Class", CType(If(row.TxnClass, DBNull.Value), Object))
                    cmd.Parameters.AddWithValue("@Amount", row.Amount)
                    cmd.ExecuteNonQuery()
                End Using
            Next
        End Using
    End Sub

    Private Function GetString(reader As OdbcDataReader, columnName As String) As String
        Dim ordinal = reader.GetOrdinal(columnName)
        Return If(reader.IsDBNull(ordinal), String.Empty, reader.GetValue(ordinal).ToString())
    End Function

    Private Function GetDateTime(reader As OdbcDataReader, columnName As String) As DateTime?
        Dim ordinal = reader.GetOrdinal(columnName)
        If reader.IsDBNull(ordinal) Then Return Nothing
        Return reader.GetDateTime(ordinal)
    End Function

    Private Function GetDecimal(reader As OdbcDataReader, columnName As String) As Decimal
        Dim ordinal = reader.GetOrdinal(columnName)
        If reader.IsDBNull(ordinal) Then Return 0D
        Return Convert.ToDecimal(reader.GetValue(ordinal))
    End Function

End Module
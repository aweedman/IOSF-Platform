Imports Microsoft.Data.SqlClient
Imports System.Data.Odbc

''' <summary>
''' Re-pulls QuickBooks' Income by Customer Detail report for a date range and replaces
''' matching rows in SQL Server - a full delete-and-reinsert refresh, not a delta.
'''
''' Table name: the SQL Server destination is IncomeByCustomerDetail_QB (not a simple
''' name-guess - confirmed against the real schema).
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
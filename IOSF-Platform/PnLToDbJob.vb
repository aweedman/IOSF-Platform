Imports Microsoft.Data.SqlClient
Imports System.Data.Odbc

''' <summary>
''' Rebuilds the PnL table for a date range from QuickBooks' Profit & Loss by Class and
''' Balance Sheet Detail reports.
'''
''' HIGH-STAKES JOB - financial data with hardcoded QuickBooks account IDs and a
''' hard-abort validation rule, both preserved exactly rather than made more forgiving:
'''  - An "Unclassified" nonzero amount aborts the entire run immediately - not logged
'''    and continued like most other jobs in this app. Whatever months were already
'''    inserted before the abort stay inserted; there's no rollback.
'''  - Once the main processing loop begins, any error aborts the whole run - not the
'''    per-row Resume-Next resilience pattern used elsewhere. The two DELETE statements
'''    at the very start are the exception: a failure there is logged but doesn't stop
'''    the run.
'''
''' CROSS-DATABASE TRANSFER: fetches every account from QuickBooks (via QODBC) into
''' memory first, then inserts into SQL Server's Account_QB table as a separate step,
''' since the two live in genuinely different databases. Account_QB is entirely deleted
''' and rebuilt on every run.
'''
''' THE sp_report SYNTAX (QuickBooks' own reporting engine, invoked via a pass-through
''' query) is used only here, and its dynamically-named Amount_N_Title/Amount_N columns
''' are read by name from the ODBC result set. This is complex, less-common territory -
''' worth running a single month first and checking the resulting P&L numbers by hand
''' against QuickBooks' own report before trusting a full multi-month run.
'''
''' HARDCODED ACCOUNT IDS AND SIGN LOGIC are specific QuickBooks internal identifiers that
''' can't be derived or guessed:
'''  - "80000159-1573175352" (Deferred Rent, old) - Amount NEGATED on insert.
'''  - "800001BA-1737144916" and "800001B9-1737144897" (Deferred Rent) - Amount as-is.
'''  - "800000B3-1476029610" and "80000054-1475519135" (Customer Security Deposits) -
'''    Amount as-is.
'''
''' PnL's own columns: Period (date), Office (the "location"/Class value), ListID (the
''' AccountListID value), Amount.
'''
''' Date range: both dates are snapped to whole-month boundaries regardless of what's
''' picked (from-date -> 1st of its month, to-date -> last day of its month).
''' </summary>
Public Module PnLToDbJob

    Private Const DeferredRentOldAccountId As String = "80000159-1573175352"
    Private Const DeferredRentAccountId1 As String = "800001BA-1737144916"
    Private Const DeferredRentAccountId2 As String = "800001B9-1737144897"
    Private Const CustomerDepositAccountId1 As String = "800000B3-1476029610"
    Private Const CustomerDepositAccountId2 As String = "80000054-1475519135"

    Public Function RunAsync(fromDate As Date, toDate As Date) As Task(Of Integer)
        Return Task.Run(Function()
                            ' Snapped to whole-month boundaries regardless of what was picked.
                            Dim monthStart = New Date(fromDate.Year, fromDate.Month, 1)
                            Dim monthEnd = New Date(toDate.Year, toDate.Month, 1).AddMonths(1).AddDays(-1)
                            Dim months = ((monthEnd.Year - monthStart.Year) * 12) + (monthEnd.Month - monthStart.Month) + 1

                            ' Phase 1: the two upfront deletes - failures here are logged but don't stop the run.
                            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                                conn.Open()
                                Try
                                    Using cmd As New SqlCommand("DELETE FROM Account_QB", conn)
                                        cmd.ExecuteNonQuery()
                                    End Using
                                Catch ex As Exception
                                    ErrorLogHelper.LogError("Upload PnL", $"SQL error in: DELETE FROM Account_QB - {ex.Message}")
                                End Try

                                Try
                                    Using cmd As New SqlCommand("DELETE FROM PnL WHERE Period BETWEEN @FromDate AND @ToDate", conn)
                                        cmd.Parameters.AddWithValue("@FromDate", monthStart)
                                        cmd.Parameters.AddWithValue("@ToDate", monthEnd)
                                        cmd.ExecuteNonQuery()
                                    End Using
                                Catch ex As Exception
                                    ErrorLogHelper.LogError("Upload PnL", $"SQL error in: DELETE FROM PnL - {ex.Message}")
                                End Try
                            End Using

                            ' Phase 2: everything from here on is a hard abort on any error - not the
                            ' per-row Resume-Next resilience pattern used elsewhere in this app.
                            Try
                                Using sqlConn As New SqlConnection(ConfigHelper.ConnectionString)
                                    sqlConn.Open()

                                    Using qbConn As New OdbcConnection(ConfigHelper.QodbcConnectionString)
                                        qbConn.Open()

                                        SyncAccountsFromQb(qbConn, sqlConn)

                                        For j = 0 To months - 1
                                            Dim thisMonthStart = monthStart.AddMonths(j)
                                            Dim thisMonthEnd = thisMonthStart.AddMonths(1).AddDays(-1)

                                            Dim aborted = ProcessProfitAndLoss(qbConn, sqlConn, thisMonthStart, thisMonthEnd)
                                            If aborted Then Return 1 ' immediate stop on an Unclassified nonzero amount

                                            ProcessBalanceSheetAdjustments(qbConn, sqlConn, thisMonthStart, thisMonthEnd)
                                        Next
                                    End Using
                                End Using
                            Catch ex As Exception
                                ErrorLogHelper.LogError("Upload PnL", $"SQL error in: {ex.Message}")
                                Return 1
                            End Try

                            Return 0
                        End Function)
    End Function

    ''' <summary>Fetches every account from QuickBooks via QODBC into memory, then inserts each one into SQL Server's Account_QB.</summary>
    Private Sub SyncAccountsFromQb(qbConn As OdbcConnection, sqlConn As SqlConnection)
        Dim accounts As New List(Of (ListId As String, Name As String, ParentRefListId As String, AccountType As String))

        Using cmd As New OdbcCommand("SELECT ListID, Name, ParentRefListID, AccountType FROM Account", qbConn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    accounts.Add((
                        reader.GetString(0),
                        reader.GetString(1),
                        If(reader.IsDBNull(2), Nothing, reader.GetString(2)),
                        If(reader.IsDBNull(3), Nothing, reader.GetString(3))
                    ))
                End While
            End Using
        End Using

        For Each acct In accounts
            Using cmd As New SqlCommand("INSERT INTO Account_QB (ListID, Name, ParentRefListID, AccountType) VALUES (@ListId, @Name, @ParentRefListId, @AccountType)", sqlConn)
                cmd.Parameters.AddWithValue("@ListId", acct.ListId)
                cmd.Parameters.AddWithValue("@Name", acct.Name)
                cmd.Parameters.AddWithValue("@ParentRefListId", If(acct.ParentRefListId, CObj(DBNull.Value)))
                cmd.Parameters.AddWithValue("@AccountType", If(acct.AccountType, CObj(DBNull.Value)))
                cmd.ExecuteNonQuery()
            End Using
        Next
    End Sub

    ''' <summary>
    ''' Runs the ProfitAndLossByClass report for one month, inserting one PnL row per
    ''' non-TOTAL, nonzero (Class, Amount) pair. Returns True if the run should abort
    ''' immediately (an Unclassified nonzero amount was found) - whatever was already
    ''' inserted for earlier months stays inserted.
    ''' </summary>
    Private Function ProcessProfitAndLoss(qbConn As OdbcConnection, sqlConn As SqlConnection, monthStart As Date, monthEnd As Date) As Boolean
        Dim fromStr = monthStart.ToString("yyyy-MM-dd")
        Dim toStr = monthEnd.ToString("yyyy-MM-dd")

        Dim reportSql = $"sp_report ProfitAndLossByClass show Amount_Title, AccountListID, Amount " &
            $"parameters ReportBasis = 'Cash', DateFrom = {{d'{fromStr}'}}, DateTo = {{d'{toStr}'}}, " &
            $"SummarizeColumnsBy = 'Class' where RowType = 'DataRow'"

        Using cmd As New OdbcCommand(reportSql, qbConn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim accountListId = reader("AccountListID").ToString()
                    Dim count = Convert.ToInt32(reader("Amount_Count"))

                    For i = 1 To count
                        Dim location = reader($"Amount_{i}_Title").ToString()
                        Dim amount = Convert.ToDecimal(reader($"Amount_{i}"))

                        If location = "Unclassified" AndAlso amount <> 0 Then
                            MessageBox.Show("Process abended. Unclassified location issue.", "Upload PnL")
                            Return True
                        ElseIf location <> "TOTAL" AndAlso amount <> 0 Then
                            InsertPnLRow(sqlConn, monthStart, location, accountListId, amount)
                        End If
                    Next
                End While
            End Using
        End Using

        Return False
    End Function

    ''' <summary>
    ''' Runs the BalanceSheetDetail report for one month, restricted to the 5 hardcoded
    ''' deferred-rent/customer-deposit accounts, grouped and summed by (AccountListID,
    ''' Class), with the sign-flip for the old Deferred Rent account - see class remarks.
    ''' </summary>
    Private Sub ProcessBalanceSheetAdjustments(qbConn As OdbcConnection, sqlConn As SqlConnection, monthStart As Date, monthEnd As Date)
        Dim fromStr = monthStart.ToString("yyyy-MM-dd")
        Dim toStr = monthEnd.ToString("yyyy-MM-dd")

        Dim reportSql = $"sp_report BalanceSheetDetail show AccountListID, Amount, Class parameters ReportBasis = 'Cash', DateFrom = {{d'{fromStr}'}}, " &
            $"DateTo = {{d'{toStr}'}} WHERE (AccountListID = '{DeferredRentOldAccountId}' OR AccountListID = '{CustomerDepositAccountId1}' OR AccountListID = '{CustomerDepositAccountId2}' " &
            $"OR AccountListID = '{DeferredRentAccountId1}' OR AccountListID = '{DeferredRentAccountId2}') " &
            $"AND Class IS NOT NULL"

        ' Grouped/summed in memory rather than via a second query against a staged
        ' intermediate table - the report output is read directly here.
        Dim grouped As New Dictionary(Of (AccountListId As String, Class_ As String), Decimal)

        Using cmd As New OdbcCommand(reportSql, qbConn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim accountListId = reader("AccountListID").ToString()
                    Dim amount = Convert.ToDecimal(reader("Amount"))
                    Dim class_ = If(reader("Class") Is DBNull.Value, Nothing, reader("Class").ToString())
                    If String.IsNullOrEmpty(class_) Then Continue While ' matches the "Class IS NOT NULL" filter above

                    Dim key = (accountListId, class_)
                    grouped(key) = grouped.GetValueOrDefault(key, 0D) + amount
                End While
            End Using
        End Using

        For Each kv In grouped
            Dim accountListId = kv.Key.AccountListId
            Dim class_ = kv.Key.Class_
            Dim amount = kv.Value
            If amount = 0 Then Continue For

            If accountListId = DeferredRentOldAccountId Then
                InsertPnLRow(sqlConn, monthStart, class_, accountListId, amount * -1) ' negated - see class remarks
            ElseIf accountListId = DeferredRentAccountId1 OrElse accountListId = DeferredRentAccountId2 Then
                InsertPnLRow(sqlConn, monthStart, class_, accountListId, amount)
            ElseIf accountListId = CustomerDepositAccountId1 OrElse accountListId = CustomerDepositAccountId2 Then
                InsertPnLRow(sqlConn, monthStart, class_, accountListId, amount)
            End If
        Next
    End Sub

    Private Sub InsertPnLRow(sqlConn As SqlConnection, period As Date, class_ As String, accountListId As String, amount As Decimal)
        Using cmd As New SqlCommand("INSERT INTO PnL (Period, Office, ListID, Amount) VALUES (@Period, @Office, @ListId, @Amount)", sqlConn)
            cmd.Parameters.AddWithValue("@Period", period)
            cmd.Parameters.AddWithValue("@Office", class_)
            cmd.Parameters.AddWithValue("@ListId", accountListId)
            cmd.Parameters.AddWithValue("@Amount", amount)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

End Module
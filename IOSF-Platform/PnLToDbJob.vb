Imports Microsoft.Data.SqlClient
Imports System.Data.Odbc

''' <summary>
''' Direct port of Landing Page.cls: Command57_Click() ("Update P&L" / "PnL to DB").
'''
''' HIGH-STAKES JOB - financial P&L data with hardcoded QuickBooks account IDs and a
''' hard-abort validation rule. Preserved exactly rather than made more forgiving:
'''  - "Unclassified" nonzero amounts ABORT THE ENTIRE RUN immediately (MsgBox, no
'''    further processing) - not logged-and-continued like most other jobs in this port.
'''    Whatever months were already inserted before the abort stay inserted; there's no
'''    rollback, matching the original exactly.
'''  - Once the main processing loop begins, ANY error aborts the whole run (matches the
'''    original's On Error GoTo ErrorHandler for this section) - NOT the per-row
'''    Resume-Next resilience pattern used elsewhere in this port. The two DELETE
'''    statements at the very start still use their own original's On Error Resume Next
'''    equivalent (failure there is logged but doesn't stop the run), matching the
'''    original's own two-phase error handling exactly.
'''
''' CROSS-DATABASE TRANSFER: Account (QODBC/QuickBooks) -> Account_QB (SQL Server) only
''' worked in the original because Access's own query engine could transparently bridge
''' two differently-linked tables (both locally aliased "Account_QB" and
''' "Account_Sync_From_QB_SQL" in Access, but pointing at genuinely different backends) in
''' one INSERT INTO ... SELECT. That's not possible here - fetches all rows from QODBC
''' into memory first, then inserts into SQL Server as a separate step. CONFIRMED with Al:
''' the SQL Server destination table's real name is "Account_QB" (not
''' "Account_Sync_From_QB" as originally guessed from the Access alias name) - it is
''' entirely deleted and rebuilt on every run, matching the original's own behavior
''' exactly.
'''
''' THE sp_report QODBC SYNTAX (QuickBooks' built-in reporting engine, invoked via a
''' pass-through query) is NOT used anywhere else in this port and is NOT independently
''' verified to work identically via System.Data.Odbc's OdbcCommand the way it did
''' through Access's QueryDef indirection. Translated as directly as possible: the
''' "sp_report ..." string is executed directly as an OdbcCommand, and the dynamically-
''' named Amount_N_Title/Amount_N columns are read by name via OdbcDataReader, matching
''' the original's own dynamic-column approach. This is genuinely untested territory -
''' recommend running a single month first and checking the P&L numbers by hand against
''' QuickBooks' own report before trusting a full multi-month run.
'''
''' HARDCODED ACCOUNT IDS AND SIGN LOGIC preserved exactly, verbatim from the original -
''' these are specific QuickBooks internal identifiers that cannot be derived or guessed:
'''  - "80000159-1573175352" (Deferred Rent, old) - Amount NEGATED on insert.
'''  - "800001BA-1737144916" and "800001B9-1737144897" (Deferred Rent) - Amount as-is.
'''  - "800000B3-1476029610" and "80000054-1475519135" (Customer Security Deposits) -
'''    Amount as-is.
'''
''' Table name NOT independently verified: PnL_SQL -> assumed real name PnL. PnL's own
''' columns ARE confirmed with Al: Period (date), Office (varchar(20) - the "location"/
''' Class value), ListID (varchar(19) - the AccountListID value), Amount (money).
''' Account_QB (destination) and the QODBC source table (Account) are also confirmed.
'''
''' Date range: unlike most other jobs' date-range dialog, both dates here are snapped to
''' whole-month boundaries regardless of what's picked (FromDate -> 1st of its month,
''' ToDate -> last day of its month), matching the original's own DateSerial(...)
''' snapping exactly.
''' </summary>
Public Module PnLToDbJob

    ' Hardcoded QuickBooks AccountListIDs - see class remarks. Verbatim from the original,
    ' not independently verified beyond what's in the source VBA.
    Private Const DeferredRentOldAccountId As String = "80000159-1573175352"
    Private Const DeferredRentAccountId1 As String = "800001BA-1737144916"
    Private Const DeferredRentAccountId2 As String = "800001B9-1737144897"
    Private Const CustomerDepositAccountId1 As String = "800000B3-1476029610"
    Private Const CustomerDepositAccountId2 As String = "80000054-1475519135"

    Public Function RunAsync(fromDate As Date, toDate As Date) As Task(Of Integer)
        Return Task.Run(Function()
                            ' Snapped to whole-month boundaries regardless of what was picked, matching the
                            ' original's own DateSerial(...) logic exactly.
                            Dim monthStart = New Date(fromDate.Year, fromDate.Month, 1)
                            Dim monthEnd = New Date(toDate.Year, toDate.Month, 1).AddMonths(1).AddDays(-1)
                            Dim months = ((monthEnd.Year - monthStart.Year) * 12) + (monthEnd.Month - monthStart.Month) + 1

                            ' Phase 1: the two upfront deletes - failures here are logged but don't stop the
                            ' run, matching the original's own On Error Resume Next for this section.
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

                            ' Phase 2: everything from here on is a HARD ABORT on any error, matching the
                            ' original's own On Error GoTo ErrorHandler for this section - not the per-row
                            ' Resume-Next resilience pattern used elsewhere in this port.
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
                                            If aborted Then Return 1 ' matches the original's own immediate Exit Sub on an Unclassified nonzero amount

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

    ''' <summary>
    ''' Fetches Account via QODBC into memory, then inserts into SQL Server's Account_QB -
    ''' the original's single cross-database INSERT INTO ... SELECT isn't possible here
    ''' (see class remarks). Uses a parameterized query instead of the original's manual
    ''' REPLACE(Name, "'", "''") string-escaping, which is unnecessary once parameters are
    ''' used.
    ''' </summary>
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
    ''' Runs the ProfitAndLossByClass sp_report for one month, inserting one PnL row per
    ''' non-TOTAL, nonzero (Class, Amount) pair. Returns True if the run should abort
    ''' immediately (an Unclassified nonzero amount was found) - matches the original's
    ''' own hard MsgBox+Exit Sub behavior exactly, including that whatever was already
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
    ''' Runs the BalanceSheetDetail sp_report for one month, restricted to the 5 hardcoded
    ''' deferred-rent/customer-deposit accounts, grouped and summed by (AccountListID,
    ''' Class), with the sign-flip logic for the old Deferred Rent account preserved
    ''' exactly - see class remarks.
    ''' </summary>
    Private Sub ProcessBalanceSheetAdjustments(qbConn As OdbcConnection, sqlConn As SqlConnection, monthStart As Date, monthEnd As Date)
        Dim fromStr = monthStart.ToString("yyyy-MM-dd")
        Dim toStr = monthEnd.ToString("yyyy-MM-dd")

        Dim reportSql = $"sp_report BalanceSheetDetail show AccountListID, Amount, Class parameters ReportBasis = 'Cash', DateFrom = {{d'{fromStr}'}}, " &
            $"DateTo = {{d'{toStr}'}} WHERE (AccountListID = '{DeferredRentOldAccountId}' OR AccountListID = '{CustomerDepositAccountId1}' OR AccountListID = '{CustomerDepositAccountId2}' " &
            $"OR AccountListID = '{DeferredRentAccountId1}' OR AccountListID = '{DeferredRentAccountId2}') " &
            $"AND Class IS NOT NULL"

        ' Grouped/summed in memory rather than via a second SELECT ... GROUP BY against
        ' the sp_report's own result (as the original did via SP_Shell) - the report
        ' output is read directly here instead of being staged in an intermediate table.
        Dim grouped As New Dictionary(Of (AccountListId As String, Class_ As String), Decimal)

        Using cmd As New OdbcCommand(reportSql, qbConn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim accountListId = reader("AccountListID").ToString()
                    Dim amount = Convert.ToDecimal(reader("Amount"))
                    Dim class_ = If(reader("Class") Is DBNull.Value, Nothing, reader("Class").ToString())
                    If String.IsNullOrEmpty(class_) Then Continue While ' matches the original's own "Class IS NOT NULL" filter

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
                InsertPnLRow(sqlConn, monthStart, class_, accountListId, amount * -1) ' negated, per the original
            ElseIf accountListId = DeferredRentAccountId1 OrElse accountListId = DeferredRentAccountId2 Then
                InsertPnLRow(sqlConn, monthStart, class_, accountListId, amount)
            ElseIf accountListId = CustomerDepositAccountId1 OrElse accountListId = CustomerDepositAccountId2 Then
                InsertPnLRow(sqlConn, monthStart, class_, accountListId, amount)
            End If
        Next
    End Sub

    ''' <summary>Column names confirmed with Al: Period, Office (the "location"/Class value), ListID (the AccountListID value), Amount.</summary>
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
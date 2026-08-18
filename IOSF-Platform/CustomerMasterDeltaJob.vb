Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command1_Click() ("QB Customer Master to DB - Delta").
''' Pulls the full QuickBooks customer list via QODBC, compares against the last-synced
''' watermark (MAX(TimeModified)) in SQL Server, and pushes only changed/new rows.
'''
''' Changes from the VBA original:
'''  - The first two operations (clearing and repopulating Customer_QB_Temp) originally
'''    ran under "On Error Resume Next" - BEFORE the "On Error GoTo ErrorHandler" line
'''    later in the sub - meaning a QODBC fetch failure would be silently swallowed with
'''    NOTHING logged, and the rest of the routine would proceed against a stale or empty
'''    temp table. This is not reproduced: the QODBC fetch is now fully error-checked and
'''    logged like everything else.
'''  - Customer_QB_Temp was a local Access table (confirmed - no linked-table descriptor).
'''    Replaced by an in-memory List(Of QbCustomerRow) via QbCustomerRepository - no DB
'''    round-trip for staging.
'''  - The DELETE-then-INSERT pair against SQL Server now runs inside a single transaction,
'''    so a crash mid-sync can't leave SQL Server with deleted-but-not-reinserted rows.
'''    The original ran these as two independent DoCmd.RunSQL calls with no such guarantee.
'''  - Table names verified against tbldefs, not assumed: "Customer_QB" here means the
'''    QODBC-linked live QuickBooks table (Access alias "Customer_QB"); the SQL Server
'''    table synced against real name "Customer_QB" too (Access alias
'''    "Customer_Sync_From_QB_SQL") - same bare name, completely different systems. See
'''    remarks on QbCustomerRepository.
'''  - Preserves the original's behavior of doing nothing if the SQL Server table has no
'''    watermark yet (MAX(TimeModified) is empty) - that's what the "Full" variant
'''    (Command5, not yet ported) is for.
'''  - Returns an error count instead of showing MsgBox/Batch-mode messaging directly -
'''    same pattern as every other job so far; the caller decides how to surface it.
''' </summary>
Public Module CustomerMasterDeltaJob

    Public Async Function RunAsync() As Task(Of Integer)
        Dim errorCount = 0
        Dim allCustomers As List(Of QbCustomerRow)

        Try
            allCustomers = Await QbCustomerRepository.FetchAllAsync()
        Catch ex As Exception
            ErrorLogHelper.LogError("Update Customer Master", $"Error retrieving customers from QuickBooks: {ex.Message}")
            Return 1
        End Try

        Dim watermark = GetWatermark()
        If watermark Is Nothing Then
            ' Matches original: If max_sql <> "" - an empty SQL Server table means nothing
            ' to delta against, so this routine intentionally does nothing further.
            Return 0
        End If

        Dim deltaRows = allCustomers.Where(Function(c) c.TimeModified.HasValue AndAlso c.TimeModified.Value > watermark.Value).ToList()
        If deltaRows.Count = 0 Then Return 0

        Try
            ApplyDelta(deltaRows)
        Catch ex As Exception
            ErrorLogHelper.LogError("Update Customer Master", $"SQL error applying customer delta: {ex.Message}")
            errorCount += 1
        End Try

        Return errorCount
    End Function

    Private Function GetWatermark() As DateTime?
        Const sql As String = "SELECT MAX(TimeModified) FROM Customer_QB"
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                Return If(result Is Nothing OrElse result Is DBNull.Value, CType(Nothing, DateTime?), CDate(result))
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Deletes then re-inserts the delta rows in SQL Server's Customer_QB table, inside
    ''' one transaction (original ran these as two independent, non-atomic statements).
    ''' </summary>
    Private Sub ApplyDelta(deltaRows As List(Of QbCustomerRow))
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            conn.Open()
            Using transaction = conn.BeginTransaction()
                Try
                    For Each row In deltaRows
                        Using deleteCmd As New SqlCommand("DELETE FROM Customer_QB WHERE AccountNumber = @AccountNumber", conn, transaction)
                            deleteCmd.Parameters.AddWithValue("@AccountNumber", row.AccountNumber)
                            deleteCmd.ExecuteNonQuery()
                        End Using

                        Const insertSql As String =
                            "INSERT INTO Customer_QB (ListID, TimeCreated, TimeModified, Name, FullName, IsActive, CompanyName, " &
                            "FirstName, ShipAddressAddr1, ShipAddressCity, ShipAddressState, Email, CustomerTypeRefFullName, " &
                            "AccountNumber, PreferredPaymentMethodRefFullName, CustomFieldTerminationInProcess) " &
                            "VALUES (@ListID, @TimeCreated, @TimeModified, @Name, @FullName, @IsActive, @CompanyName, " &
                            "@FirstName, @ShipAddressAddr1, @ShipAddressCity, @ShipAddressState, @Email, @CustomerTypeRefFullName, " &
                            "@AccountNumber, @PreferredPaymentMethodRefFullName, @CustomFieldTerminationInProcess)"

                        Using insertCmd As New SqlCommand(insertSql, conn, transaction)
                            insertCmd.Parameters.AddWithValue("@ListID", row.ListId)
                            insertCmd.Parameters.AddWithValue("@TimeCreated", CType(If(row.TimeCreated.HasValue, row.TimeCreated.Value, DBNull.Value), Object))
                            insertCmd.Parameters.AddWithValue("@TimeModified", CType(If(row.TimeModified.HasValue, row.TimeModified.Value, DBNull.Value), Object))
                            insertCmd.Parameters.AddWithValue("@Name", row.Name)
                            insertCmd.Parameters.AddWithValue("@FullName", row.FullName)
                            insertCmd.Parameters.AddWithValue("@IsActive", row.IsActive)
                            insertCmd.Parameters.AddWithValue("@CompanyName", row.CompanyName)
                            insertCmd.Parameters.AddWithValue("@FirstName", row.FirstName)
                            insertCmd.Parameters.AddWithValue("@ShipAddressAddr1", row.ShipAddressAddr1)
                            insertCmd.Parameters.AddWithValue("@ShipAddressCity", row.ShipAddressCity)
                            insertCmd.Parameters.AddWithValue("@ShipAddressState", row.ShipAddressState)
                            insertCmd.Parameters.AddWithValue("@Email", row.Email)
                            insertCmd.Parameters.AddWithValue("@CustomerTypeRefFullName", row.CustomerTypeRefFullName)
                            insertCmd.Parameters.AddWithValue("@AccountNumber", row.AccountNumber)
                            insertCmd.Parameters.AddWithValue("@PreferredPaymentMethodRefFullName", row.PreferredPaymentMethodRefFullName)
                            insertCmd.Parameters.AddWithValue("@CustomFieldTerminationInProcess", row.CustomFieldTerminationInProcess)
                            insertCmd.ExecuteNonQuery()
                        End Using
                    Next

                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

End Module
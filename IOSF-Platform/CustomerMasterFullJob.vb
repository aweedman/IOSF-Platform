Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command5_Click() ("QB Customer Master to DB - Full").
''' Same as CustomerMasterDeltaJob, per Al, except: truncates Customer_QB (SQL Server)
''' first, then loads every customer from QuickBooks - no watermark comparison, no delta
''' filtering, no per-row delete (the table is already empty after the truncate).
'''
''' Shares QbCustomerRepository/QbCustomerRow with CustomerMasterDeltaJob - same fetch
''' logic, same table-name remarks (see that file's own class comment for the
''' Customer_QB-means-two-different-things-in-two-different-systems explanation).
'''
''' TRUNCATE TABLE requires no foreign keys reference Customer_QB - not independently
''' verified. If this ever fails with a foreign-key-related error, the fix is switching
''' to DELETE FROM Customer_QB instead (slower, but respects FKs) rather than assuming
''' TRUNCATE is safe.
''' </summary>
Public Module CustomerMasterFullJob

    Public Async Function RunAsync() As Task(Of Integer)
        Dim allCustomers As List(Of QbCustomerRow)

        Try
            allCustomers = Await QbCustomerRepository.FetchAllAsync()
        Catch ex As Exception
            ErrorLogHelper.LogError("Update Customer Master (Full)", $"Error retrieving customers from QuickBooks: {ex.Message}")
            Return 1
        End Try

        Try
            ReloadAll(allCustomers)
        Catch ex As Exception
            ErrorLogHelper.LogError("Update Customer Master (Full)", $"SQL error reloading customer master: {ex.Message}")
            Return 1
        End Try

        Return 0
    End Function

    ''' <summary>Truncates Customer_QB, then inserts every fetched customer - all inside one transaction, same atomicity guarantee as the Delta job's own ApplyDelta.</summary>
    Private Sub ReloadAll(customers As List(Of QbCustomerRow))
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            conn.Open()
            Using transaction = conn.BeginTransaction()
                Try
                    Using truncateCmd As New SqlCommand("TRUNCATE TABLE Customer_QB", conn, transaction)
                        truncateCmd.ExecuteNonQuery()
                    End Using

                    Const insertSql As String =
                        "INSERT INTO Customer_QB (ListID, TimeCreated, TimeModified, Name, FullName, IsActive, CompanyName, " &
                        "FirstName, ShipAddressAddr1, ShipAddressCity, ShipAddressState, Email, CustomerTypeRefFullName, " &
                        "AccountNumber, PreferredPaymentMethodRefFullName, CustomFieldTerminationInProcess) " &
                        "VALUES (@ListID, @TimeCreated, @TimeModified, @Name, @FullName, @IsActive, @CompanyName, " &
                        "@FirstName, @ShipAddressAddr1, @ShipAddressCity, @ShipAddressState, @Email, @CustomerTypeRefFullName, " &
                        "@AccountNumber, @PreferredPaymentMethodRefFullName, @CustomFieldTerminationInProcess)"

                    For Each row In customers
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
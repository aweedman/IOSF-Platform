Imports System.Data.Common
Imports System.Data.Odbc

''' <summary>
''' Fetches every row from QuickBooks' Customer table via QODBC (Access alias "Customer_QB",
''' real QuickBooks table "Customer" - confirmed via its tbldefs .json, same pattern as
''' the Invoice_QB lookup earlier). Shared by CustomerMasterDeltaJob and the not-yet-ported
''' full-sync job, since both need the identical data.
''' </summary>
Public Module QbCustomerRepository

    Public Function FetchAllAsync() As Task(Of List(Of QbCustomerRow))
        Return Task.Run(Function()
                            Dim result As New List(Of QbCustomerRow)
                            Const sql As String =
                                "SELECT ListID, TimeCreated, TimeModified, Name, FullName, IsActive, CompanyName, FirstName, " &
                                "ShipAddressAddr1, ShipAddressCity, ShipAddressState, Email, CustomerTypeRefFullName, " &
                                "AccountNumber, PreferredPaymentMethodRefFullName, CustomFieldTerminationInProcess " &
                                "FROM Customer"

                            Using conn As New OdbcConnection(ConfigHelper.QodbcConnectionString)
                                Using cmd As New OdbcCommand(sql, conn)
                                    conn.Open()
                                    Using reader = cmd.ExecuteReader()
                                        While reader.Read()
                                            result.Add(New QbCustomerRow With {
                                                .ListId = GetString(reader, "ListID"),
                                                .TimeCreated = GetDateTime(reader, "TimeCreated"),
                                                .TimeModified = GetDateTime(reader, "TimeModified"),
                                                .Name = GetString(reader, "Name"),
                                                .FullName = GetString(reader, "FullName"),
                                                .IsActive = GetString(reader, "IsActive") = "1" OrElse GetString(reader, "IsActive").Equals("True", StringComparison.OrdinalIgnoreCase),
                                                .CompanyName = GetString(reader, "CompanyName"),
                                                .FirstName = GetString(reader, "FirstName"),
                                                .ShipAddressAddr1 = GetString(reader, "ShipAddressAddr1"),
                                                .ShipAddressCity = GetString(reader, "ShipAddressCity"),
                                                .ShipAddressState = GetString(reader, "ShipAddressState"),
                                                .Email = GetString(reader, "Email"),
                                                .CustomerTypeRefFullName = GetString(reader, "CustomerTypeRefFullName"),
                                                .AccountNumber = GetString(reader, "AccountNumber"),
                                                .PreferredPaymentMethodRefFullName = GetString(reader, "PreferredPaymentMethodRefFullName"),
                                                .CustomFieldTerminationInProcess = GetString(reader, "CustomFieldTerminationInProcess")
                                            })
                                        End While
                                    End Using
                                End Using
                            End Using

                            Return result
                        End Function)
    End Function

    Private Function GetString(reader As OdbcDataReader, columnName As String) As String
        Dim ordinal = reader.GetOrdinal(columnName)
        Return If(reader.IsDBNull(ordinal), String.Empty, reader.GetValue(ordinal).ToString())
    End Function

    Private Function GetDateTime(reader As OdbcDataReader, columnName As String) As DateTime?
        Dim ordinal = reader.GetOrdinal(columnName)
        If reader.IsDBNull(ordinal) Then Return Nothing
        Return reader.GetDateTime(ordinal)
    End Function

End Module
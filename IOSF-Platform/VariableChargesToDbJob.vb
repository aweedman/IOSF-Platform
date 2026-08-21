Imports Microsoft.Data.SqlClient

''' <summary>
''' Syncs variable (usage-based) charges for a date range from HostedSuite into
''' Variable_Charges.
'''
''' Field mapping from the API's ChargeInfo: ServiceName (the human-readable service name)
''' is stored in the Service column. EntityStatus is available on the API response but not
''' stored - it isn't used anywhere downstream.
'''
''' Pagination starts at Page=0. Results are deduplicated by Id (a HashSet tracks seen
''' Ids while paging) as a safety net against the API ever returning overlapping pages.
'''
''' DELETE + all INSERTs for the date range run inside a single transaction, so a
''' mid-loop failure can't leave the range half-synced - either the whole range applies
''' cleanly or none of it does.
''' </summary>
Public Module VariableChargesToDbJob

    Private Const ApiBaseUrl As String = "https://io2.hostedsuite.com/api/"
    Private Const PageSize As Integer = 1000

    Public Async Function RunAsync(startDate As Date, endDate As Date) As Task(Of Integer)
        Dim errorCount = 0
        Dim charges As New List(Of ChargeInfo)

        Try
            charges = Await FetchChargesAsync(startDate, endDate)
        Catch ex As Exception
            ErrorLogHelper.LogError("Variable Charges to DB", $"Error retrieving charges: {ex.Message}")
            Return 1
        End Try

        Try
            ApplyVariableCharges(startDate, endDate, charges)
        Catch ex As Exception
            ErrorLogHelper.LogError("Variable Charges to DB", $"SQL error applying variable charges: {ex.Message}")
            errorCount += 1
        End Try

        Return errorCount
    End Function

    Private Async Function FetchChargesAsync(startDate As Date, endDate As Date) As Task(Of List(Of ChargeInfo))
        Dim result As New List(Of ChargeInfo)
        Dim seenIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Dim rangeStart = $"{startDate.Month}/{startDate.Day}/{startDate.Year} 00:00"
        Dim rangeEnd = $"{endDate.Month}/{endDate.Day}/{endDate.Year} 23:59"
        Dim dateRange = $"{{Start:{rangeStart},End:{rangeEnd}}}"

        Dim headers = New Dictionary(Of String, String) From {{"Authorization", HostedSuiteAuth.ComputeAuthHeader()}}
        Dim page = 0
        Dim totalPages = 1

        While page < totalPages
            Dim queryParams = New Dictionary(Of String, String) From {
                {"DateOfCharge", dateRange},
                {"Page", page.ToString()},
                {"CountPerPage", PageSize.ToString()}
            }

            Dim response = Await ApiClient.GetAsync($"{ApiBaseUrl}charges", queryParams, headers, timeoutSeconds:=30)
            response.EnsureSuccess()

            Dim data = response.DataAs(Of ChargesListResponse)()
            If data.Items IsNot Nothing Then
                For Each item In data.Items
                    If seenIds.Add(item.Id) Then result.Add(item)
                Next
            End If
            totalPages = Math.Max(data.TotalPages, 1)
            page += 1
        End While

        Return result
    End Function

    Private Sub ApplyVariableCharges(startDate As Date, endDate As Date, charges As List(Of ChargeInfo))
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            conn.Open()
            Using transaction = conn.BeginTransaction()
                Try
                    Using deleteCmd As New SqlCommand("DELETE FROM Variable_Charges WHERE TransactionDate BETWEEN @StartDate AND @EndDate", conn, transaction)
                        deleteCmd.Parameters.AddWithValue("@StartDate", startDate)
                        deleteCmd.Parameters.AddWithValue("@EndDate", endDate)
                        deleteCmd.ExecuteNonQuery()
                    End Using

                    For Each charge In charges
                        Dim quantity = If(charge.Quantity.GetValueOrDefault() = 0, 1, charge.Quantity.GetValueOrDefault())
                        Dim transactionDate = DateTime.Parse(charge.DateOfCharge).Date

                        Const insertSql As String =
                            "INSERT INTO Variable_Charges (Id, ClientId, Company_Evo, Service, TransactionDate, Qty, Cost, Description) " &
                            "VALUES (@Id, @ClientId, @CompanyEvo, @Service, @TransactionDate, @Qty, @Cost, @Description)"

                        Using cmd As New SqlCommand(insertSql, conn, transaction)
                            cmd.Parameters.AddWithValue("@Id", charge.Id)
                            cmd.Parameters.AddWithValue("@ClientId", charge.ClientId)
                            cmd.Parameters.AddWithValue("@CompanyEvo", If(charge.ClientName, String.Empty))
                            cmd.Parameters.AddWithValue("@Service", If(charge.ServiceName, String.Empty))
                            cmd.Parameters.AddWithValue("@TransactionDate", transactionDate)
                            cmd.Parameters.AddWithValue("@Qty", quantity)
                            cmd.Parameters.AddWithValue("@Cost", charge.Cost.GetValueOrDefault())
                            cmd.Parameters.AddWithValue("@Description", If(charge.Description, String.Empty))
                            cmd.ExecuteNonQuery()
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
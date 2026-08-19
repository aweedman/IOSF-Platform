Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command70_Click() ("Staff Assisted Scan Extra Pages
''' to Evo").
'''
''' CUT OVER TO THE NEW API, same as CopierChargesToEvoJob - POSTs to
''' io2.hostedsuite.com/api/charges (HostedSuiteAuth Authorization header) instead of the
''' older io.hostedsuite.com/api/json/reply/ NewChargeRequest endpoint (plain-body
''' credentials, no auth header). Same REAL BUG FIX as that job too: the original never
''' checked whether either charge POST actually succeeded - this checks
''' response.EnsureSuccess() per charge and logs a real error if it fails.
'''
''' Main aggregation query computes, per ClientId: TotQty = SUM(Qty) - (COUNT(Id) * 10) -
''' the pages beyond the first 10 on each ">10 page" scan job, summed across all such jobs
''' for that client - and Cnt = COUNT(Id), the number of such jobs. Two SEPARATE charges
''' are posted per client: one for the job COUNT (ServiceId 525c64110f4e161c8025c1ab), one
''' for the ADDITIONAL PAGES total (ServiceId 6a2b1a8cbf94090b3ce75254) - preserved exactly,
''' not combined into one charge.
'''
''' Service name string literals preserved EXACTLY, character-for-character, including an
''' inconsistency in the original itself: "Scanning Less or = 10 Pages - Staff Assisted  (
''' # of scans )" has a DOUBLE space before the opening parenthesis, while "Scanning
''' Greater 10 Pages - Staff Assisted ( total # pages )" has a single space. Not "cleaned
''' up" - these need to match whatever is actually stored in Variable_Charges.Service.
'''
''' Validation warnings (NOT charge postings - Error_Log notifications for human review)
''' preserved exactly:
'''  - Any Company_Evo with a "<=10 pages" scan job where Qty > 6 (unusually high for a
'''    supposedly-small job).
'''  - Any Company_Evo with a ">10 pages" scan job where Qty < 11 (unusually low for a
'''    supposedly-large job).
'''  - A single, non-per-client "Other Charge Warning" if ANY Variable_Charges row in the
'''    date range has Service = 'Other' - matches the original's DLookup-based existence
'''    check exactly (one warning total, not one per matching row).
'''
''' Table name confirmed: Variable_Charges (from Variable_Charges_SQL) was already
''' established as a real table/schema in VariableChargesToDbJob earlier in this port -
''' not re-guessed here.
'''
''' Per-row/per-query failures are logged and do NOT stop the rest, matching the
''' original's own On Error Resume Next.
''' </summary>
Public Module ScanExtraPagesToEvoJob

    Private Const ApiBaseUrl As String = "https://io2.hostedsuite.com/api"

    Private Const ServiceId_ScanCount As String = "525c64110f4e161c8025c1ab"
    Private Const ServiceId_AdditionalPages As String = "6a2b1a8cbf94090b3ce75254"

    Private Const ServiceName_Greater10 As String = "Scanning Greater 10 Pages - Staff Assisted ( total # pages )"
    Private Const ServiceName_Less10 As String = "Scanning Less or = 10 Pages - Staff Assisted  ( # of scans )" ' double space before "(" - preserved exactly, see class remarks

    Public Async Function RunAsync(billStartDate As Date, billEndDate As Date, postingDate As Date) As Task(Of Integer)
        Dim errorCount = 0

        Dim rows As New List(Of (ClientId As String, TotQty As Integer, Cnt As Integer))
        Try
            Const sql As String =
                "SELECT ClientId, (SUM(Qty) - (COUNT(Id) * 10)) AS TotQty, COUNT(Id) AS Cnt " &
                "FROM Variable_Charges " &
                "WHERE TransactionDate >= @BillStart AND TransactionDate <= @BillEnd AND Service = @ServiceName " &
                "GROUP BY ClientId"

            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@BillStart", billStartDate)
                    cmd.Parameters.AddWithValue("@BillEnd", billEndDate)
                    cmd.Parameters.AddWithValue("@ServiceName", ServiceName_Greater10)
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            rows.Add((reader.GetValue(0).ToString(), Convert.ToInt32(reader.GetValue(1)), Convert.ToInt32(reader.GetValue(2))))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ErrorLogHelper.LogError("Staff Assisted Extra Charges to Evo", $"SQL error: {ex.Message}")
            Return 1
        End Try

        Dim token = HostedSuiteAuth.ComputeAuthHeader()
        Dim headers = New Dictionary(Of String, String) From {{"Authorization", token}}

        For Each row In rows
            Try
                Await PostCharge(headers, postingDate, ServiceId_ScanCount, row.ClientId, row.Cnt)
            Catch ex As Exception
                ErrorLogHelper.LogError("Staff Assisted Extra Charges to Evo", $"Error posting scan-count charge for ClientId {row.ClientId}: {ex.Message}")
                errorCount += 1
            End Try

            Try
                Await PostCharge(headers, postingDate, ServiceId_AdditionalPages, row.ClientId, row.TotQty)
            Catch ex As Exception
                ErrorLogHelper.LogError("Staff Assisted Extra Charges to Evo", $"Error posting additional-pages charge for ClientId {row.ClientId}: {ex.Message}")
                errorCount += 1
            End Try
        Next

        errorCount += LogQtyWarnings(billStartDate, billEndDate, ServiceName_Less10, "Qty > 6", "Large value for Scans <= 10 Pages for  ")
        errorCount += LogQtyWarnings(billStartDate, billEndDate, ServiceName_Greater10, "Qty < 11", "Small value for Scans > 10 Pages for  ")

        If HasOtherCharge(billStartDate, billEndDate) Then
            ErrorLogHelper.LogError("Staff Assisted Extra Pages", "Other Charge Warning")
            errorCount += 1
        End If

        Return errorCount
    End Function

    Private Async Function PostCharge(headers As Dictionary(Of String, String), postingDate As Date, serviceId As String, clientId As String, quantity As Integer) As Task
        Dim payload = New With {
            .dateOfCharge = postingDate.ToString("yyyy-MM-dd"),
            .serviceId = serviceId,
            .clientId = clientId,
            .quantity = quantity,
            .notes = "BillingCycle"
        }
        Dim response = Await ApiClient.PostAsync($"{ApiBaseUrl}/charges", payload, headers, timeoutSeconds:=60)
        response.EnsureSuccess()
    End Function

    ''' <summary>
    ''' qtyCondition is embedded directly since it's always one of two fixed, known-safe
    ''' literals from this file's own two call sites, not user input.
    ''' </summary>
    Private Function LogQtyWarnings(billStartDate As Date, billEndDate As Date, serviceName As String, qtyCondition As String, warningPrefix As String) As Integer
        Dim errorCount = 0
        Try
            Dim sql = $"SELECT DISTINCT Company_Evo FROM Variable_Charges " &
                      $"WHERE TransactionDate >= @BillStart AND TransactionDate <= @BillEnd AND {qtyCondition} AND Service = @ServiceName"

            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@BillStart", billStartDate)
                    cmd.Parameters.AddWithValue("@BillEnd", billEndDate)
                    cmd.Parameters.AddWithValue("@ServiceName", serviceName)
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim companyEvo = If(reader.IsDBNull(0), "", reader.GetString(0))
                            ErrorLogHelper.LogError("Variable Charges to QB", $"{warningPrefix}{companyEvo}")
                            errorCount += 1
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ErrorLogHelper.LogError("Staff Assisted Extra Charges to Evo", $"SQL error checking Qty warnings for {serviceName}: {ex.Message}")
            errorCount += 1
        End Try
        Return errorCount
    End Function

    Private Function HasOtherCharge(billStartDate As Date, billEndDate As Date) As Boolean
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand("SELECT TOP 1 Id FROM Variable_Charges WHERE Service = 'Other' AND TransactionDate >= @BillStart AND TransactionDate <= @BillEnd", conn)
                cmd.Parameters.AddWithValue("@BillStart", billStartDate)
                cmd.Parameters.AddWithValue("@BillEnd", billEndDate)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                Return result IsNot Nothing AndAlso result IsNot DBNull.Value
            End Using
        End Using
    End Function

End Module
Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command69_Click() ("Copier Charges to Evo").
'''
''' CUT OVER TO THE NEW API per Al: the original POSTed to the older
''' io.hostedsuite.com/api/json/reply/ NewChargeRequest endpoint, with credentials sent as
''' plain body parameters (CustomerName/UserName/Password - no auth header at all). This
''' instead POSTs to the newer io2.hostedsuite.com/api/charges endpoint (same
''' Authorization-header auth already established via HostedSuiteAuth for the GET variant
''' of this same endpoint in VariableChargesToDbJob/SpheremailChargesToDbJob), confirmed
''' against that endpoint's own metadata (io2.hostedsuite.com/api/json/metadata?op=NewCharge)
''' before building this.
'''
''' REAL BUG FIXED: the original never checked whether each NewChargeRequest call actually
''' succeeded - it fired the request and moved on regardless, with no way to know a charge
''' silently failed to post (network issue, auth failure, etc.). This checks
''' response.EnsureSuccess() per charge and logs a real error if it fails, rather than
''' silently continuing.
'''
''' Three near-identical query+post loops (Black & White, Color, Scan) factored into one
''' shared PostChargesForQuery helper, differing only in the SQL query and the hardcoded
''' ServiceId. The three queries' filter differences are preserved exactly as written in
''' the original, not homogenized: BW uses "HAVING SUM(...) > 0" (filtering the aggregate),
''' Color uses "WHERE total_color_pages > 0" (filtering rows before aggregation - though
''' functionally equivalent here, since summing in zero-value rows doesn't change a sum),
''' and Scan has no "printed = 'Y'" condition at all (scans apparently have no such
''' concept, matching the original's own omission for that one query specifically).
'''
''' The Posting Date (DateOfCharge sent to the API) is a SEPARATE date from the billing
''' cycle start/end used to filter Printer_Usage_Log - matches the original exactly,
''' where InvDate is a third, independently-prompted date, not derived from the other two.
'''
''' Hardcoded ServiceIds preserved verbatim from the original - these are specific
''' HostedSuite internal identifiers that cannot be derived or guessed:
'''  - "6a2b0cefbf94090b3c605950" - Print/Copy Black & White
'''  - "6a2b1127bf94090b3c89e615" - Print/Copy Color
'''  - "6a2b12e3bf94090b3c98f97d" - Scan
'''
''' Evo ClientId lookup is an EXACT match on Evo_Customer_XRef.ThirdPartyAccountID,
''' matching the original's own DLookup exactly (unlike KubeMeetingsToDbJob's prefix
''' match against a different field, Name - these are genuinely different lookups in
''' different jobs, not inconsistent with each other).
'''
''' Table name NOT independently verified: Printer_Usage_Log_SQL -> assumed real name
''' Printer_Usage_Log (the simple-strip convention used elsewhere in this port, but not
''' confirmed against a tbldefs descriptor). Customer_Ops_Item and Evo_Customer_XRef are
''' both confirmed real tables elsewhere in this port already.
'''
''' Per-row failures (both the ClientId lookup and the charge POST) are logged and do NOT
''' stop the rest, matching the original's own On Error Resume Next.
''' </summary>
Public Module CopierChargesToEvoJob

    Private Const ApiBaseUrl As String = "https://io2.hostedsuite.com/api"

    Private Const ServiceId_BlackAndWhite As String = "6a2b0cefbf94090b3c605950"
    Private Const ServiceId_Color As String = "6a2b1127bf94090b3c89e615"
    Private Const ServiceId_Scan As String = "6a2b12e3bf94090b3c98f97d"

    Public Async Function RunAsync(billStartDate As Date, billEndDate As Date, postingDate As Date) As Task(Of Integer)
        Dim errorCount = 0

        Dim bwSql = $"
            SELECT Customer_Ops_Item.Account_Num, SUM(Printer_Usage_Log.total_pages - Printer_Usage_Log.total_color_pages) AS Qty
            FROM Printer_Usage_Log
            INNER JOIN Customer_Ops_Item ON Printer_Usage_Log.user_name = Customer_Ops_Item.Fac_Code
            WHERE Printer_Usage_Log.usage_day >= @BillStart AND Printer_Usage_Log.usage_day <= @BillEnd
            AND (Printer_Usage_Log.job_type = 'PRINT' OR Printer_Usage_Log.job_type = 'COPY')
            AND Printer_Usage_Log.cancelled = 'N' AND Printer_Usage_Log.printed = 'Y' AND Printer_Usage_Log.refunded = 'N'
            GROUP BY Customer_Ops_Item.Account_Num
            HAVING SUM(Printer_Usage_Log.total_pages - Printer_Usage_Log.total_color_pages) > 0"

        Dim colorSql = $"
            SELECT Customer_Ops_Item.Account_Num, SUM(Printer_Usage_Log.total_color_pages) AS Qty
            FROM Printer_Usage_Log
            INNER JOIN Customer_Ops_Item ON Printer_Usage_Log.user_name = Customer_Ops_Item.Fac_Code
            WHERE Printer_Usage_Log.usage_day >= @BillStart AND Printer_Usage_Log.usage_day <= @BillEnd
            AND (Printer_Usage_Log.job_type = 'PRINT' OR Printer_Usage_Log.job_type = 'COPY')
            AND Printer_Usage_Log.cancelled = 'N' AND Printer_Usage_Log.printed = 'Y' AND Printer_Usage_Log.refunded = 'N'
            AND Printer_Usage_Log.total_color_pages > 0
            GROUP BY Customer_Ops_Item.Account_Num"

        Dim scanSql = $"
            SELECT Customer_Ops_Item.Account_Num, SUM(Printer_Usage_Log.total_pages) AS Qty
            FROM Printer_Usage_Log
            INNER JOIN Customer_Ops_Item ON Printer_Usage_Log.user_name = Customer_Ops_Item.Fac_Code
            WHERE Printer_Usage_Log.usage_day >= @BillStart AND Printer_Usage_Log.usage_day <= @BillEnd
            AND Printer_Usage_Log.job_type = 'SCAN'
            AND Printer_Usage_Log.cancelled = 'N' AND Printer_Usage_Log.refunded = 'N'
            GROUP BY Customer_Ops_Item.Account_Num"

        errorCount += Await PostChargesForQuery(bwSql, billStartDate, billEndDate, ServiceId_BlackAndWhite, postingDate)
        errorCount += Await PostChargesForQuery(colorSql, billStartDate, billEndDate, ServiceId_Color, postingDate)
        errorCount += Await PostChargesForQuery(scanSql, billStartDate, billEndDate, ServiceId_Scan, postingDate)

        Return errorCount
    End Function

    Private Async Function PostChargesForQuery(sql As String, billStartDate As Date, billEndDate As Date, serviceId As String, postingDate As Date) As Task(Of Integer)
        Dim errorCount = 0
        Dim rows As New List(Of (AccountNum As String, Qty As Double))

        Try
            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@BillStart", billStartDate)
                    cmd.Parameters.AddWithValue("@BillEnd", billEndDate)
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            rows.Add((reader.GetValue(0).ToString(), Convert.ToDouble(reader.GetValue(1))))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ErrorLogHelper.LogError("Copier Charges to Evo", $"SQL error in: {sql} - {ex.Message}")
            Return 1
        End Try

        Dim token = HostedSuiteAuth.ComputeAuthHeader()
        Dim headers = New Dictionary(Of String, String) From {{"Authorization", token}}

        For Each row In rows
            Try
                Dim clientId = LookupEvoClientId(row.AccountNum)
                If String.IsNullOrEmpty(clientId) Then
                    ErrorLogHelper.LogError("Copier Charges to Evo", $"Evo ClientId not found for Account_Num {row.AccountNum}")
                    errorCount += 1
                    Continue For
                End If

                Dim payload = New With {
                    .dateOfCharge = postingDate.ToString("yyyy-MM-dd"),
                    .serviceId = serviceId,
                    .clientId = clientId,
                    .quantity = row.Qty,
                    .notes = "BillingCycle"
                }

                Dim response = Await ApiClient.PostAsync($"{ApiBaseUrl}/charges", payload, headers, timeoutSeconds:=60)
                response.EnsureSuccess()
            Catch ex As Exception
                ErrorLogHelper.LogError("Copier Charges to Evo", $"Error posting charge for Account_Num {row.AccountNum}: {ex.Message}")
                errorCount += 1
            End Try
        Next

        Return errorCount
    End Function

    Private Function LookupEvoClientId(accountNum As String) As String
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand("SELECT Id FROM Evo_Customer_XRef WHERE ThirdPartyAccountId = @AccountNum", conn)
                cmd.Parameters.AddWithValue("@AccountNum", accountNum)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                Return If(result Is Nothing OrElse result Is DBNull.Value, Nothing, result.ToString())
            End Using
        End Using
    End Function

End Module
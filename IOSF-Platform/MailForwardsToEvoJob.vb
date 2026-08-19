Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command68_Click() ("Mail Forwards to Evo").
'''
''' CUT OVER TO THE NEW API, same as CopierChargesToEvoJob/ScanExtraPagesToEvoJob - POSTs
''' to io2.hostedsuite.com/api/charges (HostedSuiteAuth Authorization header) instead of
''' the older io.hostedsuite.com/api/json/reply/ NewChargeRequest endpoint. Same REAL BUG
''' FIX as those jobs: the original never checked whether either charge POST actually
''' succeeded - this checks response.EnsureSuccess() per charge and logs a real error if
''' it fails.
'''
''' REAL BUG FIXED (found and confirmed by tracing the original's own logic, not
''' guessed): the "subtract prior removals" step - DSum("Qty", "Variable_Charges_SQL",
''' "ClientID = '...' AND Service = 'Mail Forward ( Remove )'") - returns Null when a
''' client has ZERO such charges on record. The original's very next line, "If Subtract >
''' 0", would then compare Null > 0, which VBA raises as a runtime error rather than
''' evaluating to False. Since the whole procedure runs under On Error Resume Next, that
''' error gets silently logged as a generic SQL error, and Resume Next then skips the
''' REST of that loop iteration - including the charge-posting code that comes after it.
''' Net effect: any client with no prior "removal" charges on record would likely never
''' get their mail-forward charge posted at all. Fixed here by treating a no-rows DSum
''' result as 0, not an error - this is NOT reproduced, since it doesn't read as
''' deliberate behavior.
'''
''' MARKUP FORMULA extracted into named constants per Al, for easy future updates:
''' MarkupPercentage (0.2 = 20%), MarkupCap (3.0 = $3 max surcharge per shipment),
''' RoundingIncrement (0.05 = round to nearest nickel). Formula preserved exactly:
''' per-shipment marked-up cost = ROUND((BaseCost + MIN(BaseCost * MarkupPercentage,
''' MarkupCap)) / RoundingIncrement) * RoundingIncrement, summed across all of an
''' account's shipments, then rounded to 2 decimals. Kept as SQL (with the constants
''' passed in as parameters) rather than reimplemented in VB.NET, since the rounding
''' happens PER SHIPMENT before summing - moving this to VB.NET would mean fetching every
''' individual shipment row instead of using SQL's own GROUP BY aggregation, a much
''' larger structural change for no real benefit.
'''
''' Two-source UNION preserved exactly: USPS forwards (SendPro directly, filtered by
''' Carrier='USPS' or a specific Carrier_Acct 'RY6026', excluding Voided/Refunded, within
''' the billing date range) UNIONed with FedEx forwards (SendPro INNER JOINed to FedEx on
''' Tracking_Num, filtered by FedEx.Billing_Start_Date = EXACT billing start date - not a
''' range - matching the original's own "=" comparison, not ">="/"<=" like the USPS side).
'''
''' Account_Num < 9000 threshold uses a NUMERIC comparison (unquoted in the original VBA),
''' consistent with SendPro.Account_Num apparently being numeric there - NOT
''' independently verified against a real schema query, unlike Customer_QB.AccountNumber
''' (confirmed nvarchar(99) elsewhere in this port) or Customer_QB's own account-number
''' filtering (also confirmed numeric-is-fine by Al for Customer Master's gallery).
'''
''' UCASE(...)/LIKE '*...*' (Access syntax) translated to UPPER(...)/LIKE '%...%' (T-SQL)
''' - same values/logic, different wildcard/function syntax only.
'''
''' Hardcoded ServiceIds preserved verbatim: "69f4e87ba11f931ee4851416" (forward count),
''' "6a2afc83bf94090b3cad7a71" (forwarding amount/cost).
'''
''' Table names confirmed: SendPro, FedEx, Variable_Charges, Evo_Customer_XRef all already
''' established as real tables/schemas elsewhere in this port - not re-guessed here.
'''
''' Per-row failures are logged and do NOT stop the rest, matching the original's own On
''' Error Resume Next (aside from the Null-comparison bug fixed above, which was never
''' meant to stop anything in the first place).
''' </summary>
Public Module MailForwardsToEvoJob

    Private Const ApiBaseUrl As String = "https://io2.hostedsuite.com/api"

    Private Const ServiceId_ForwardCount As String = "69f4e87ba11f931ee4851416"
    Private Const ServiceId_ForwardAmount As String = "6a2afc83bf94090b3cad7a71"

    ' Markup formula constants moved to the shared MailForwardMarkup module, per Al, so
    ' this job and MailForwardsReportJob can't drift apart and either can be updated in
    ' one place. See MailForwardMarkup.vb.

    Private Const AccountNumberThreshold As Integer = 9000

    Public Async Function RunAsync(billStartDate As Date, billEndDate As Date, postingDate As Date) As Task(Of Integer)
        Dim errorCount = 0

        errorCount += Await PostForwardCounts(billStartDate, billEndDate, postingDate)
        errorCount += Await PostForwardAmounts(billStartDate, billEndDate, postingDate)

        Return errorCount
    End Function

    Private Async Function PostForwardCounts(billStartDate As Date, billEndDate As Date, postingDate As Date) As Task(Of Integer)
        Dim errorCount = 0
        Dim rows As New List(Of (AccountNum As String, Quantity As Integer))

        Const sql As String = "
            SELECT Account_Num, SUM(Quantity) AS Total_Quantity
            FROM (
                SELECT Account_Num, COUNT(Account_Num) AS Quantity
                FROM SendPro
                WHERE Transaction_Date >= @BillStart AND Transaction_Date <= @BillEnd
                AND SM_Status <> 'Voided' AND SM_Status NOT LIKE '%Refund%'
                AND (UPPER(Carrier) = 'USPS' OR Carrier_Acct = 'RY6026') AND Account_Num < @AccountThreshold
                GROUP BY Account_Num

                UNION ALL

                SELECT SendPro.Account_Num, COUNT(SendPro.Account_Num) AS Quantity
                FROM SendPro
                INNER JOIN FedEx ON SendPro.Tracking_Num = FedEx.Tracking_Num
                WHERE FedEx.Billing_Start_Date = @BillStart AND SendPro.Account_Num < @AccountThreshold
                GROUP BY SendPro.Account_Num
            ) AS Union_Table
            GROUP BY Account_Num"

        Try
            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@BillStart", billStartDate)
                    cmd.Parameters.AddWithValue("@BillEnd", billEndDate)
                    cmd.Parameters.AddWithValue("@AccountThreshold", AccountNumberThreshold)
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            rows.Add((reader.GetValue(0).ToString(), Convert.ToInt32(reader.GetValue(1))))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ErrorLogHelper.LogError("Mail Forwards to Evo", $"SQL error: {ex.Message}")
            Return 1
        End Try

        Dim headers = New Dictionary(Of String, String) From {{"Authorization", HostedSuiteAuth.ComputeAuthHeader()}}

        For Each row In rows
            Try
                Dim clientId = LookupEvoClientId(row.AccountNum)
                If String.IsNullOrEmpty(clientId) Then
                    ErrorLogHelper.LogError("Mail Forwards to Evo", $"Evo ClientId not found for Account_Num {row.AccountNum}")
                    errorCount += 1
                    Continue For
                End If

                Dim subtract = GetPriorRemovalQty(clientId)
                Dim quantity = row.Quantity
                If subtract > 0 Then quantity -= subtract

                ' REAL ISSUE FIXED: confirmed via a real 500 "Quantity is required" error
                ' from the new API - a zero or negative quantity (which can genuinely
                ' happen here if a client's prior removals equal or exceed this cycle's
                ' new forwards) is rejected outright. The original never guarded against
                ' this, possibly because the older API it originally called was less
                ' strict. Skipped rather than attempted, since there's nothing meaningful
                ' to charge for in that case anyway.
                If quantity <= 0 Then
                    ErrorLogHelper.LogError("Mail Forwards to Evo", $"Skipped forward-count charge for Account_Num {row.AccountNum}: computed quantity is {quantity} after subtracting prior removals ({subtract})")
                    Continue For
                End If

                Dim payload = New With {
                    .dateOfCharge = postingDate.ToString("yyyy-MM-dd"),
                    .serviceId = ServiceId_ForwardCount,
                    .clientId = clientId,
                    .quantity = quantity,
                    .notes = "BillingCycle"
                }
                Try
                    Dim response = Await ApiClient.PostAsync($"{ApiBaseUrl}/charges", payload, headers, timeoutSeconds:=60)
                    response.EnsureSuccess()
                Catch ex As Exception
                    ' TEMPORARY DIAGNOSTIC - includes the exact outgoing JSON payload
                    ' alongside the error, so the request can be directly compared
                    ' against what the server said back, rather than guessing at
                    ' serialization again.
                    Dim sentJson = Text.Json.JsonSerializer.Serialize(payload)
                    ErrorLogHelper.LogError("Mail Forwards to Evo", $"Error posting forward-count charge for Account_Num {row.AccountNum}: {ex.Message} | Sent: {sentJson}")
                    errorCount += 1
                    Continue For
                End Try
            Catch ex As Exception
                ErrorLogHelper.LogError("Mail Forwards to Evo", $"Error preparing forward-count charge for Account_Num {row.AccountNum}: {ex.Message}")
                errorCount += 1
            End Try
        Next

        Return errorCount
    End Function

    Private Async Function PostForwardAmounts(billStartDate As Date, billEndDate As Date, postingDate As Date) As Task(Of Integer)
        Dim errorCount = 0
        Dim rows As New List(Of (AccountNum As String, Amount As Decimal))

        Const sql As String = "
            SELECT Account_Num, ROUND(SUM(Amount), 2) AS Total_Amount
            FROM (
                SELECT Account_Num,
                    ROUND((Total_Cost + IIF(Total_Cost * @MarkupPct < @MarkupCap, ROUND(Total_Cost * @MarkupPct, 2), @MarkupCap)) / @RoundIncrement, 0) * @RoundIncrement AS Amount
                FROM SendPro
                WHERE Transaction_Date >= @BillStart AND Transaction_Date <= @BillEnd
                AND SM_Status <> 'Voided' AND SM_Status NOT LIKE '%Refund%'
                AND (UPPER(Carrier) = 'USPS' OR Carrier_Acct = 'RY6026') AND Account_Num < @AccountThreshold

                UNION ALL

                SELECT SendPro.Account_Num,
                    ROUND((FedEx.Total_Cost + IIF(FedEx.Total_Cost * @MarkupPct < @MarkupCap, ROUND(FedEx.Total_Cost * @MarkupPct, 2), @MarkupCap)) / @RoundIncrement, 0) * @RoundIncrement AS Amount
                FROM SendPro
                INNER JOIN FedEx ON SendPro.Tracking_Num = FedEx.Tracking_Num
                WHERE FedEx.Billing_Start_Date = @BillStart AND SendPro.Account_Num < @AccountThreshold
            ) AS Union_Table
            GROUP BY Account_Num"

        Try
            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@BillStart", billStartDate)
                    cmd.Parameters.AddWithValue("@BillEnd", billEndDate)
                    cmd.Parameters.AddWithValue("@AccountThreshold", AccountNumberThreshold)
                    cmd.Parameters.AddWithValue("@MarkupPct", MailForwardMarkup.MarkupPercentage)
                    cmd.Parameters.AddWithValue("@MarkupCap", MailForwardMarkup.MarkupCap)
                    cmd.Parameters.AddWithValue("@RoundIncrement", MailForwardMarkup.RoundingIncrement)
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            rows.Add((reader.GetValue(0).ToString(), Convert.ToDecimal(reader.GetValue(1))))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ErrorLogHelper.LogError("Mail Forwards to Evo", $"SQL error: {ex.Message}")
            Return 1
        End Try

        Dim headers = New Dictionary(Of String, String) From {{"Authorization", HostedSuiteAuth.ComputeAuthHeader()}}

        For Each row In rows
            Try
                Dim clientId = LookupEvoClientId(row.AccountNum)
                If String.IsNullOrEmpty(clientId) Then
                    ErrorLogHelper.LogError("Mail Forwards to Evo", $"Evo ClientId not found for Account_Num {row.AccountNum}")
                    errorCount += 1
                    Continue For
                End If

                Dim payload = New With {
                    .dateOfCharge = postingDate.ToString("yyyy-MM-dd"),
                    .serviceId = ServiceId_ForwardAmount,
                    .clientId = clientId,
                    .quantity = 1,
                    .cost = row.Amount,
                    .notes = "BillingCycle"
                }
                Dim response = Await ApiClient.PostAsync($"{ApiBaseUrl}/charges", payload, headers, timeoutSeconds:=60)
                response.EnsureSuccess()
            Catch ex As Exception
                ErrorLogHelper.LogError("Mail Forwards to Evo", $"Error posting forward-amount charge for Account_Num {row.AccountNum}: {ex.Message}")
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

    ''' <summary>
    ''' No date filter, matching the original's own DSum exactly - sums ALL-time "Mail
    ''' Forward ( Remove )" charges for this client, not just within the current billing
    ''' cycle. Returns 0 (not an error) when no matching rows exist - see class remarks
    ''' for the real bug this fixes.
    ''' </summary>
    Private Function GetPriorRemovalQty(clientId As String) As Decimal
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand("SELECT SUM(Qty) FROM Variable_Charges WHERE ClientId = @ClientId AND Service = 'Mail Forward ( Remove )'", conn)
                cmd.Parameters.AddWithValue("@ClientId", clientId)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                Return If(result Is Nothing OrElse result Is DBNull.Value, 0D, Convert.ToDecimal(result))
            End Using
        End Using
    End Function

End Module
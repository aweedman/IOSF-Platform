Imports Microsoft.Data.SqlClient

''' <summary>
''' Posts mail forwarding charges to HostedSuite for a billing cycle: a per-account
''' forward count and a per-account forwarding amount (postage + markup).
'''
''' MARKUP FORMULA constants (MarkupPercentage, MarkupCap, RoundingIncrement) live in the
''' shared MailForwardMarkup module, so this job and MailForwardsReportJob can't drift
''' apart - either can be updated in one place. Formula: per-shipment marked-up cost =
''' ROUND((BaseCost + MIN(BaseCost * MarkupPercentage, MarkupCap)) / RoundingIncrement) *
''' RoundingIncrement, summed across all of an account's shipments, then rounded to 2
''' decimals. This stays in SQL (with the constants passed in as parameters) rather than
''' being computed in VB.NET, since the rounding happens per shipment before summing -
''' doing that in VB.NET would mean fetching every individual shipment row instead of
''' using SQL's own GROUP BY aggregation.
'''
''' Two-source UNION: USPS forwards (SendPro directly, filtered by Carrier='USPS' or a
''' specific Carrier_Acct 'RY6026', excluding Voided/Refunded, within the billing date
''' range) combined with FedEx forwards (SendPro joined to FedEx on Tracking_Num, filtered
''' by FedEx.Billing_Start_Date matching the billing start date exactly - not a range,
''' unlike the USPS side).
'''
''' GetPriorRemovalQty returns 0 (not a database NULL) when a client has no prior "Mail
''' Forward ( Remove )" charges on record - this matters because the caller subtracts this
''' value from the forward count, and treating "no rows" as 0 rather than propagating a
''' NULL/error keeps that subtraction well-defined for every client, including ones who've
''' never had a removal charge at all.
'''
''' If subtracting prior removals brings a client's forward count to zero or below (their
''' removals equal or exceed this cycle's new forwards), that charge is skipped entirely
''' rather than posted - HostedSuite rejects a zero or negative quantity outright, and
''' there's nothing meaningful to charge for in that case anyway.
'''
''' Hardcoded ServiceIds are HostedSuite's own internal identifiers and can't be derived:
''' "69f4e87ba11f931ee4851416" (forward count), "6a2afc83bf94090b3cad7a71" (forwarding
''' amount/cost).
''' </summary>
Public Module MailForwardsToEvoJob

    Private Const ApiBaseUrl As String = "https://io2.hostedsuite.com/api"

    Private Const ServiceId_ForwardCount As String = "69f4e87ba11f931ee4851416"
    Private Const ServiceId_ForwardAmount As String = "6a2afc83bf94090b3cad7a71"

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
                Dim response = Await ApiClient.PostAsync($"{ApiBaseUrl}/charges", payload, headers, timeoutSeconds:=60)
                response.EnsureSuccess()
            Catch ex As Exception
                ErrorLogHelper.LogError("Mail Forwards to Evo", $"Error posting forward-count charge for Account_Num {row.AccountNum}: {ex.Message}")
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

    ''' <summary>Sums ALL-time "Mail Forward ( Remove )" charges for this client, not just within the current billing cycle. Returns 0 when no matching rows exist.</summary>
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
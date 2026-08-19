Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command71_Click() ("Spheremail to Evo").
'''
''' CUT OVER TO THE NEW API, same as the other "...to Evo" jobs in this port - POSTs to
''' io2.hostedsuite.com/api/charges (HostedSuiteAuth Authorization header) instead of the
''' older io.hostedsuite.com/api/json/reply/ NewChargeRequest endpoint. Same REAL BUG FIX
''' as those jobs: the original never checked whether any charge POST actually succeeded -
''' this checks response.EnsureSuccess() per charge and logs a real error if it fails.
'''
''' Customer_Header_Join is a NEW table/view name not seen elsewhere in this port - based
''' on the columns referenced (Mail_Box, AccountNumber, IsActive), this looks like a
''' pre-joined view combining Customer_Ops_Header.Mail_Box with Customer_QB.AccountNumber/
''' IsActive, giving a convenient Mail_Box -> AccountNumber mapping (SphereMail identifies
''' customers by mailbox; Evo/HostedSuite by account number). NOT independently verified -
''' queried directly by the columns the original source references, trusting it exists
''' with that exact shape.
'''
''' SUBTLE BEHAVIOR PRESERVED EXACTLY: the "No PMB match found" error-logging path (when
''' AccountNumber comes back blank from the LEFT JOIN) is copy-pasted into all three
''' query loops, but is only PRACTICALLY reachable for the Mail Pickup query. The Scan and
''' Envelope Picture queries filter WHERE Customer_Header_Join.IsActive = 1 - since a
''' LEFT JOIN with no match leaves IsActive NULL, and NULL = 1 is NULL (not TRUE) in both
''' Access and SQL Server, any no-match row is already excluded by the WHERE clause before
''' it could ever trigger the "No PMB match" check for those two. Mail Pickup has no such
''' IsActive filter, so no-match rows DO reach it there. This looks like copy-pasted
''' shared code across three near-identical blocks (a pattern already seen elsewhere in
''' this port) rather than a deliberate difference - preserved as-is rather than "cleaned
''' up", since removing the dead-for-two-of-three-queries check could introduce a subtle
''' difference if this SQL-null-semantics reasoning is even slightly off.
'''
''' Mail Pickup query structure preserved exactly: counts DISTINCT (Mail_Box,
''' AccountNumber, Txn_Date) combinations via a subquery, then counts Txn_Date per group -
''' NOT a plain COUNT(Id) like the other two queries. This looks like a deliberate choice
''' to count distinct pickup DAYS rather than raw Spheremail_Charges rows (which could
''' have multiple rows per pickup event/day for some reason) - not simplified to match
''' the other two queries' simpler COUNT(Id) pattern.
'''
''' Hardcoded ServiceIds preserved verbatim:
'''  - "525c64110f4e161c8025c1ab" (Scan count) - the SAME ServiceId already used in
'''    ScanExtraPagesToEvoJob for its own "scan count" charge - makes sense, since both
'''    represent a scan-count billable unit under the same Evo service category, just
'''    from different source systems (copier/PaperCut scans vs SphereMail mailroom scans).
'''  - "6a2b2752bf94090b3c5c29fb" (Envelope Picture)
'''  - "4f7f35b10117b11bc8264b1e" (Mail Pickup)
'''
''' Table names NOT independently verified: Spheremail_Charges_SQL -> assumed real name
''' Spheremail_Charges (already used, same assumption, in SpheremailChargesToDbJob
''' earlier in this port). Customer_Header_Join_SQL -> assumed real name
''' Customer_Header_Join (new to this job, not confirmed elsewhere). Evo_Customer_XRef is
''' confirmed real elsewhere in this port already.
'''
''' Per-row failures are logged and do NOT stop the rest, matching the original's own On
''' Error Resume Next.
''' </summary>
Public Module SpheremailToEvoJob

    Private Const ApiBaseUrl As String = "https://io2.hostedsuite.com/api"

    Private Const ServiceId_Scan As String = "525c64110f4e161c8025c1ab"
    Private Const ServiceId_EnvelopePicture As String = "6a2b2752bf94090b3c5c29fb"
    Private Const ServiceId_MailPickup As String = "4f7f35b10117b11bc8264b1e"

    Public Async Function RunAsync(billStartDate As Date, billEndDate As Date, postingDate As Date) As Task(Of Integer)
        Dim errorCount = 0

        Const scanSql As String = "
            SELECT Spheremail_Charges.Mail_Box, Customer_Header_Join.AccountNumber, COUNT(Spheremail_Charges.Id) AS Qty
            FROM Spheremail_Charges
            LEFT JOIN Customer_Header_Join ON Spheremail_Charges.Mail_Box = Customer_Header_Join.Mail_Box
            WHERE Customer_Header_Join.IsActive = 1 AND Spheremail_Charges.Txn_Date >= @BillStart AND Spheremail_Charges.Txn_Date <= @BillEnd
            AND Spheremail_Charges.Description = 'Scan'
            GROUP BY Spheremail_Charges.Mail_Box, Customer_Header_Join.AccountNumber"

        Const envelopePictureSql As String = "
            SELECT Spheremail_Charges.Mail_Box, Customer_Header_Join.AccountNumber, COUNT(Spheremail_Charges.Id) AS Qty
            FROM Spheremail_Charges
            LEFT JOIN Customer_Header_Join ON Spheremail_Charges.Mail_Box = Customer_Header_Join.Mail_Box
            WHERE Customer_Header_Join.IsActive = 1 AND Spheremail_Charges.Txn_Date >= @BillStart AND Spheremail_Charges.Txn_Date <= @BillEnd
            AND Spheremail_Charges.Description = 'Mail Item Picture'
            GROUP BY Spheremail_Charges.Mail_Box, Customer_Header_Join.AccountNumber"

        ' No IsActive filter here, unlike the two queries above - see class remarks.
        Const mailPickupSql As String = "
            SELECT z.Mail_Box, z.AccountNumber, COUNT(z.Txn_Date) AS Qty
            FROM (
                SELECT DISTINCT Spheremail_Charges.Mail_Box, Customer_Header_Join.AccountNumber, Spheremail_Charges.Txn_Date
                FROM Spheremail_Charges
                LEFT JOIN Customer_Header_Join ON Spheremail_Charges.Mail_Box = Customer_Header_Join.Mail_Box
                WHERE Spheremail_Charges.Txn_Date >= @BillStart AND Spheremail_Charges.Txn_Date <= @BillEnd
                AND Spheremail_Charges.Description = 'Mail Item Pickup'
            ) AS z
            GROUP BY z.Mail_Box, z.AccountNumber"

        errorCount += Await PostChargesForQuery(scanSql, billStartDate, billEndDate, ServiceId_Scan, postingDate)
        errorCount += Await PostChargesForQuery(envelopePictureSql, billStartDate, billEndDate, ServiceId_EnvelopePicture, postingDate)
        errorCount += Await PostChargesForQuery(mailPickupSql, billStartDate, billEndDate, ServiceId_MailPickup, postingDate)

        Return errorCount
    End Function

    Private Async Function PostChargesForQuery(sql As String, billStartDate As Date, billEndDate As Date, serviceId As String, postingDate As Date) As Task(Of Integer)
        Dim errorCount = 0
        Dim rows As New List(Of (MailBox As String, AccountNum As String, Qty As Integer))

        Try
            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@BillStart", billStartDate)
                    cmd.Parameters.AddWithValue("@BillEnd", billEndDate)
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim mailBox = If(reader.IsDBNull(0), "", reader.GetValue(0).ToString())
                            Dim accountNum = If(reader.IsDBNull(1), "", reader.GetValue(1).ToString())
                            Dim qty = Convert.ToInt32(reader.GetValue(2))
                            rows.Add((mailBox, accountNum, qty))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ErrorLogHelper.LogError("Spheremail charges to Evo", $"SQL error: {ex.Message}")
            Return 1
        End Try

        Dim headers = New Dictionary(Of String, String) From {{"Authorization", HostedSuiteAuth.ComputeAuthHeader()}}

        For Each row In rows
            If String.IsNullOrEmpty(row.AccountNum) Then
                ErrorLogHelper.LogError("Speheremail Charges to Evo", $"No PMB match found for {row.MailBox}") ' "Speheremail" typo preserved verbatim from the original's own Error_Log source string
                errorCount += 1
                Continue For
            End If

            Try
                Dim clientId = LookupEvoClientId(row.AccountNum)
                If String.IsNullOrEmpty(clientId) Then
                    ErrorLogHelper.LogError("Spheremail charges to Evo", $"Evo ClientId not found for Account_Num {row.AccountNum}")
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
                ErrorLogHelper.LogError("Spheremail charges to Evo", $"Error posting charge for Account_Num {row.AccountNum}: {ex.Message}")
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
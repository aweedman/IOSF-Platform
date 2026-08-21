Imports Microsoft.Data.SqlClient

''' <summary>
''' Posts Spheremail mailroom charges (Scan, Envelope Picture, Mail Pickup) to
''' HostedSuite for a billing cycle.
'''
''' Customer_Header_Join is a view mapping Mail_Box to AccountNumber/IsActive (Spheremail
''' identifies customers by mailbox number; HostedSuite by account number), joined against
''' Spheremail_Charges to translate mailbox-based charges into account-based ones.
'''
''' NOTE ON A SUBTLE SQL BEHAVIOR: the "No PMB match found" warning (logged when a mail
''' item's AccountNumber comes back blank from the LEFT JOIN) is only reachable for the
''' Mail Pickup query, not Scan or Envelope Picture. Those two filter on
''' "Customer_Header_Join.IsActive = 1" - and since a LEFT JOIN with no match leaves
''' IsActive NULL, and "NULL = 1" evaluates to NULL (not TRUE) in SQL, any no-match row is
''' already excluded by that WHERE clause before it could reach the warning check. Mail
''' Pickup has no such IsActive filter, so no-match rows do reach it there. The warning
''' check is left in all three query loops rather than removed from the two where it can
''' never fire, in case that reasoning turns out to be wrong in some edge case.
'''
''' Mail Pickup's query counts DISTINCT (Mail_Box, AccountNumber, Txn_Date) combinations
''' via a subquery, rather than a plain row count like the other two - this counts
''' distinct pickup days, not raw Spheremail_Charges rows (which can have more than one
''' row per pickup event/day).
'''
''' Hardcoded ServiceIds are HostedSuite's own internal identifiers and can't be derived:
'''  - "525c64110f4e161c8025c1ab" (Scan count) - the same ServiceId used for the copier/
'''    PaperCut scan-count charge elsewhere in this app, since both represent the same
'''    billable "scan count" category, just from different source systems.
'''  - "6a2b2752bf94090b3c5c29fb" (Envelope Picture)
'''  - "4f7f35b10117b11bc8264b1e" (Mail Pickup)
'''
''' Each charge post and each SQL check is independent - a failure in one doesn't stop the
''' rest, and every failure is logged.
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
                ErrorLogHelper.LogError("Speheremail Charges to Evo", $"No PMB match found for {row.MailBox}") ' matches the exact string already used elsewhere in Error_Log for this warning - not a typo to "fix"
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
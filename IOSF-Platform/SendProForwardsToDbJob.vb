Imports Microsoft.Data.SqlClient
Imports CsvHelper
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms

''' <summary>
''' Imports a SendPro mail-forwarding CSV export into the SendPro table, resolving each
''' recipient/sender to an internal account number.
'''
''' Account lookups (Customer_QB by Name, SendPro_XRef by Company) are pre-fetched into
''' in-memory dictionaries once, rather than queried per CSV row - the file can be large,
''' so this avoids up to four separate SQL queries per row. Both dictionaries use
''' case-insensitive keys. If a table has duplicate Name/Company values, the last one
''' encountered wins.
'''
''' Rows where no account number could be resolved (falls back to the placeholder account
''' number "1") are tallied and reported once via a single end-of-run popup ("Data import
''' complete. N records without account number. Update in DB") rather than logged
''' individually - mismatches are managed through a separate tool, not Error_Log.
'''
''' The returned error count reflects only genuine SQL/exception failures - the
''' no-account-match count is tracked and reported separately via the popup above, not
''' folded into the error count.
'''
''' The FCM (First-Class Mail) sequential counter is read, incremented, and persisted back
''' to Config for every applicable row (not once per run) - a number is "burned"
''' (persisted) even if that row's later INSERT fails, since the counter update happens
''' before the INSERT is attempted.
'''
''' Before processing any rows, existing SendPro rows are deleted for the Transaction_Date
''' range found in the CSV file itself (its own min/max date), making it safe to re-run
''' the same file, or an overlapping one, without creating duplicates. This isn't wrapped
''' in one transaction with the row-by-row import - the delete is its own quick upfront
''' step, and each row's failure is handled independently without rolling back the rest.
''' </summary>
Public Module SendProForwardsToDbJob

    Private Const NoAccountPlaceholder As String = "1"
    Private Const SenderCompanyToIgnore As String = "Intelligent Office"

    ' 1-indexed CSV column positions.
    Private Const C_PackageTrackingNumber As Integer = 1
    Private Const C_TransactionDate As Integer = 2
    Private Const C_TrackingStatus As Integer = 4
    Private Const C_Status As Integer = 5
    Private Const C_RecipientName As Integer = 9
    Private Const C_RecipientCompany As Integer = 10
    Private Const C_RecipientAddress As Integer = 11
    Private Const C_RecipientCountry As Integer = 12
    Private Const C_Carrier As Integer = 13
    Private Const C_CarrierAccountNumber As Integer = 14
    Private Const C_Class As Integer = 15
    Private Const C_ServiceCost As Integer = 22
    Private Const C_SenderCompany As Integer = 40
    Private Const C_TotalAdjustedCost As Integer = 46

    Public Function RunAsync(csvFilePath As String) As Task(Of Integer)
        Return Task.Run(Function()
                            Dim errorCount = 0
                            Dim noAccountCount = 0

                            Dim customerAccountByName = FetchCustomerAccountsByName()
                            Dim sendProXrefByCompany = FetchSendProXrefByCompany()

                            ' Two-phase read: load every row's raw fields into memory first, both to
                            ' find the min/max Transaction_Date (for the upfront delete) and to stay
                            ' consistent with the "stop at first blank TransactionDate" rule below,
                            ' which needs to happen before either step.
                            Dim rows = ReadCsvRows(csvFilePath)
                            If rows.Count = 0 Then Return 0

                            Dim minDate = rows.Min(Function(r) r.TransactionDate)
                            Dim maxDate = rows.Max(Function(r) r.TransactionDate)

                            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                                conn.Open()

                                Using deleteCmd As New SqlCommand("DELETE FROM SendPro WHERE Transaction_Date BETWEEN @MinDate AND @MaxDate", conn)
                                    deleteCmd.Parameters.AddWithValue("@MinDate", minDate.Date)
                                    deleteCmd.Parameters.AddWithValue("@MaxDate", maxDate.Date)
                                    deleteCmd.ExecuteNonQuery()
                                End Using

                                For Each row In rows
                                    Try
                                        If ProcessRow(conn, row, customerAccountByName, sendProXrefByCompany) Then
                                            noAccountCount += 1
                                        End If
                                    Catch ex As Exception
                                        ErrorLogHelper.LogError("Upload Mail Forwards", $"SQL error: {ex.Message}")
                                        errorCount += 1
                                    End Try
                                Next
                            End Using

                            MessageBox.Show($"Data import complete. {noAccountCount} records without account number. Update in DB", "SendPro Forwards to DB")

                            Return errorCount
                        End Function)
    End Function

    ''' <summary>Raw fields for one CSV row, captured up front so the file only needs to be read once.</summary>
    Private Class SendProCsvRow
        Public Property TrackingNumber As String
        Public Property TransactionDate As Date
        Public Property TrackingStatus As String
        Public Property Status As String
        Public Property RecipientName As String
        Public Property RecipientCompany As String
        Public Property RecipientAddress As String
        Public Property RecipientCountry As String
        Public Property Carrier As String
        Public Property CarrierAccountNumber As String
        Public Property Class_ As String
        Public Property ServiceCost As String
        Public Property SenderCompany As String
        Public Property TotalAdjustedCost As String
    End Class

    ''' <summary>Reads every row up front, stopping at the first blank Transaction_Date, since some CSV exports have trailing footer/blank rows that would otherwise fail to parse as a date.</summary>
    Private Function ReadCsvRows(csvFilePath As String) As List(Of SendProCsvRow)
        Dim result As New List(Of SendProCsvRow)

        Using reader = New StreamReader(csvFilePath)
            Using csv = New CsvReader(reader, CultureInfo.InvariantCulture)
                csv.Read() ' header row
                csv.ReadHeader()

                While csv.Read()
                    Dim rawDate = GetField(csv, C_TransactionDate)
                    If String.IsNullOrEmpty(rawDate) Then Exit While

                    result.Add(New SendProCsvRow With {
                        .TrackingNumber = GetField(csv, C_PackageTrackingNumber),
                        .TransactionDate = DateTime.Parse(rawDate),
                        .TrackingStatus = GetField(csv, C_TrackingStatus),
                        .Status = GetField(csv, C_Status),
                        .RecipientName = GetField(csv, C_RecipientName),
                        .RecipientCompany = GetField(csv, C_RecipientCompany),
                        .RecipientAddress = GetField(csv, C_RecipientAddress),
                        .RecipientCountry = GetField(csv, C_RecipientCountry),
                        .Carrier = GetField(csv, C_Carrier),
                        .CarrierAccountNumber = GetField(csv, C_CarrierAccountNumber),
                        .Class_ = GetField(csv, C_Class),
                        .ServiceCost = GetField(csv, C_ServiceCost),
                        .SenderCompany = GetField(csv, C_SenderCompany),
                        .TotalAdjustedCost = GetField(csv, C_TotalAdjustedCost)
                    })
                End While
            End Using
        End Using

        Return result
    End Function

    ''' <summary>Returns True if this row had no resolvable account number - tallied by the caller and reported once via a summary popup, not logged per-row.</summary>
    Private Function ProcessRow(conn As SqlConnection, row As SendProCsvRow,
                                 customerAccountByName As Dictionary(Of String, String),
                                 sendProXrefByCompany As Dictionary(Of String, String)) As Boolean
        Dim acct = ResolveAccountNumber(row.RecipientCompany, row.SenderCompany, customerAccountByName, sendProXrefByCompany)
        Dim noAccountFound = (acct = NoAccountPlaceholder)

        Dim trackingNum = row.TrackingNumber

        If row.Class_ = "First-Class Mail" Then
            Dim nextNum = IncrementFcmCounter(conn)
            trackingNum = $"FCM {nextNum}"
        End If

        trackingNum = trackingNum.Replace("{", "").Replace("}", "")

        Const insertSql As String =
            "INSERT INTO SendPro (Account_Num, Tracking_Num, Transaction_Date, Tracking_Status, SM_Status, " &
            "Recipient, Company, Address, Country, Carrier, Carrier_Acct, Class, Total_Cost, Service_Cost, Company_Sender) " &
            "VALUES (@AccountNum, @TrackingNum, @TransactionDate, @TrackingStatus, @SmStatus, " &
            "@Recipient, @Company, @Address, @Country, @Carrier, @CarrierAcct, @Class, @TotalCost, @ServiceCost, @CompanySender)"

        Using cmd As New SqlCommand(insertSql, conn)
            cmd.Parameters.AddWithValue("@AccountNum", acct)
            cmd.Parameters.AddWithValue("@TrackingNum", trackingNum)
            cmd.Parameters.AddWithValue("@TransactionDate", row.TransactionDate)
            cmd.Parameters.AddWithValue("@TrackingStatus", row.TrackingStatus)
            cmd.Parameters.AddWithValue("@SmStatus", row.Status)
            cmd.Parameters.AddWithValue("@Recipient", row.RecipientName)
            cmd.Parameters.AddWithValue("@Company", row.RecipientCompany)
            cmd.Parameters.AddWithValue("@Address", row.RecipientAddress)
            cmd.Parameters.AddWithValue("@Country", row.RecipientCountry)
            cmd.Parameters.AddWithValue("@Carrier", row.Carrier)
            cmd.Parameters.AddWithValue("@CarrierAcct", row.CarrierAccountNumber)
            cmd.Parameters.AddWithValue("@Class", row.Class_)
            cmd.Parameters.AddWithValue("@TotalCost", ParseCostOrZero(row.TotalAdjustedCost))
            cmd.Parameters.AddWithValue("@ServiceCost", ParseCostOrZero(row.ServiceCost))
            cmd.Parameters.AddWithValue("@CompanySender", row.SenderCompany)
            cmd.ExecuteNonQuery()
        End Using

        Return noAccountFound
    End Function

    ''' <summary>Four-tier fallback, in order: recipient company in Customer_QB, then sender company (excluding "Intelligent Office" and blank) in Customer_QB, then recipient company in SendPro_XRef, then sender company in SendPro_XRef. Falls back to the "1" placeholder if none resolve.</summary>
    Private Function ResolveAccountNumber(recipientCompany As String, senderCompany As String,
                                           customerAccountByName As Dictionary(Of String, String),
                                           sendProXrefByCompany As Dictionary(Of String, String)) As String
        Dim acct As String = Nothing

        If Not String.IsNullOrEmpty(recipientCompany) Then
            customerAccountByName.TryGetValue(recipientCompany, acct)
        End If

        If String.IsNullOrEmpty(acct) AndAlso Not String.IsNullOrEmpty(senderCompany) AndAlso senderCompany <> SenderCompanyToIgnore Then
            customerAccountByName.TryGetValue(senderCompany, acct)
        End If

        If String.IsNullOrEmpty(acct) AndAlso Not String.IsNullOrEmpty(recipientCompany) Then
            sendProXrefByCompany.TryGetValue(recipientCompany, acct)
        End If

        If String.IsNullOrEmpty(acct) AndAlso Not String.IsNullOrEmpty(senderCompany) AndAlso senderCompany <> SenderCompanyToIgnore Then
            sendProXrefByCompany.TryGetValue(senderCompany, acct)
        End If

        Return If(String.IsNullOrEmpty(acct), NoAccountPlaceholder, acct)
    End Function

    Private Function FetchCustomerAccountsByName() As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand("SELECT Name, AccountNumber FROM Customer_QB", conn)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        If Not reader.IsDBNull(0) Then result(reader.GetString(0)) = If(reader.IsDBNull(1), Nothing, Convert.ToString(reader.GetValue(1)))
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    ''' <summary>Reads via GetValue()+Convert.ToString() rather than a strict GetInt32/GetString call, since Account_Num's exact numeric type shouldn't matter here.</summary>
    Private Function FetchSendProXrefByCompany() As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand("SELECT Company, Account_Num FROM SendPro_XRef", conn)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        If Not reader.IsDBNull(0) Then result(reader.GetString(0)) = If(reader.IsDBNull(1), Nothing, Convert.ToString(reader.GetValue(1)))
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Function IncrementFcmCounter(conn As SqlConnection) As Integer
        Dim current As Integer
        Using cmd As New SqlCommand("SELECT Low FROM Config WHERE Name = 'SendPro FCM Counter'", conn)
            current = CInt(cmd.ExecuteScalar())
        End Using

        Dim next_ = current + 1
        Using cmd As New SqlCommand("UPDATE Config SET Low = @Low WHERE Name = 'SendPro FCM Counter'", conn)
            cmd.Parameters.AddWithValue("@Low", next_.ToString())
            cmd.ExecuteNonQuery()
        End Using

        Return next_
    End Function

    Private Function GetField(csv As CsvReader, oneIndexedColumn As Integer) As String
        Return If(csv.GetField(oneIndexedColumn - 1), String.Empty).Trim()
    End Function

    Private Function ParseCostOrZero(s As String) As Decimal
        Dim result As Decimal
        Decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, result)
        Return result
    End Function

End Module
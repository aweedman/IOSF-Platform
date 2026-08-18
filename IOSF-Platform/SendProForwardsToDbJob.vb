Imports Microsoft.Data.SqlClient
Imports CsvHelper
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms

''' <summary>
''' Direct port of Landing Page.cls: Command15_Click() ("160.1 - SendPro Forwards to DB").
'''
''' REQUIRES ADDING THE CsvHelper NUGET PACKAGE to the project (same compile-test caveat as
''' ClosedXML/Ical.Net elsewhere in this port - could not be verified against the real
''' library in this sandbox). Used instead of naive comma-splitting because recipient
''' address fields will very likely contain embedded commas needing proper quoted-CSV
''' handling.
'''
''' Deviations from the original (confirmed reasoning, not guessed):
'''  - The file dialog's AllowMultiSelect=True is unused in the original - only
'''    .SelectedItems(1) is ever read afterward, so multi-select had no actual effect.
'''    This uses a plain single-file picker matching the real behavior.
'''  - Account lookups (Customer_QB by Name, SendPro_XRef by Company) are pre-fetched into
'''    in-memory Dictionaries once, instead of up to 4 separate SQL queries per CSV row -
'''    same efficiency pattern already applied to KubePaymentsToQbJob's duplicate check,
'''    for the same reason (this could be a large file). Both dictionaries use
'''    case-insensitive keys, matching Access/Jet's default case-insensitive string
'''    comparison behavior for the DLookup calls this replaces. If a table genuinely has
'''    duplicate Name/Company values, this takes the LAST one encountered when building the
'''    dictionary - not necessarily identical to whichever row DLookup happened to return
'''    first, but no evidence duplicate names are expected to exist regardless.
'''  - Rows where no account number could be resolved (falls back to the placeholder
'''    Account_Num "1") are tallied and reported via a single end-of-run popup, matching
'''    the original's own summary message exactly ("Data import complete. N records
'''    without account number. Update in DB"). An earlier version of this file logged each
'''    one individually to Error_Log instead - reverted per Al, who has a separate
'''    PowerApps tool for managing these mappings and doesn't need per-row entries.
'''  - Returns an error count reflecting only genuine SQL/exception failures - the
'''    no-account-match count above is tracked and reported separately, not folded into
'''    this count, matching what "error count" means everywhere else in this port.
'''
''' Table names NOT independently verified: SendPro_XRef_SQL -> assumed real name
''' SendPro_XRef, SendPro_SQL -> assumed real name SendPro (both the simple-strip
''' convention that's held for most tables in this port, but neither confirmed against a
''' tbldefs descriptor). Customer_Sync_From_QB_SQL -> confirmed real name Customer_QB
''' (already established elsewhere in this port).
'''
''' The FCM (First-Class Mail) sequential counter is read, incremented, and persisted back
''' to Config PER ROW (not once per run) - preserved exactly, including the original's
''' behavior that a number is "burned" (persisted) even if that row's later INSERT fails,
''' since the counter update happens before the INSERT is attempted.
'''
''' DELETE-BEFORE-INSERT ADDED per Al (not in the original): before processing any rows,
''' deletes existing SendPro rows whose Transaction_Date falls within the MIN/MAX
''' Transaction_Date found in the CSV file itself - makes re-running the same file (or an
''' overlapping one) idempotent for testing/reloading, instead of creating duplicates. The
''' range is derived from the file's own data rather than a separate date-range prompt (no
''' such dialog exists for this job, unlike Call Counts/Variable Charges) - confirmed with
''' Al this is the intended source for the range. This does NOT wrap the whole run in one
''' atomic transaction the way Call Counts/Variable Charges do - the delete is its own
''' quick upfront step, and the per-row insert loop keeps its existing Resume-Next
''' behavior (a failure on one row doesn't roll back or stop the rest), matching the
''' original's own per-row resilience design, which nothing here was asked to change.
'''
''' Per-row failures are logged and do NOT stop the rest, matching the original's
''' On Error Resume Next.
''' </summary>
Public Module SendProForwardsToDbJob

    Private Const NoAccountPlaceholder As String = "1"
    Private Const SenderCompanyToIgnore As String = "Intelligent Office"

    ' 1-indexed CSV column positions, matching the original's VBA constants exactly.
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
                            ' find the min/max Transaction_Date (for the upfront delete, per Al - makes
                            ' testing/reloading the same file idempotent instead of creating duplicates)
                            ' and to stay consistent with the existing "stop at first blank
                            ' TransactionDate" behavior below, which needs to happen before either step.
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

                            ' Matches the original's summary popup exactly ("Data import complete. N
                            ' records without account number. Update in DB") instead of a per-row
                            ' Error_Log entry, per Al - he has a separate PowerApps tool for updating
                            ' these mappings and doesn't need each one logged individually.
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

    ''' <summary>
    ''' Reads every row up front, stopping at the first blank Transaction_Date - matches
    ''' the original's exact loop condition ("While TransactionDate <> """), rather than
    ''' looping to end-of-file. Some CSV exports have trailing footer/blank rows; without
    ''' this, such a row would throw on DateTime.Parse and get logged as a spurious SQL
    ''' error instead of cleanly ending the import.
    ''' </summary>
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

    ''' <summary>
    ''' Returns True if this row had no resolvable account number. No longer logged
    ''' per-row to Error_Log (was, in an earlier version of this file) - per Al, he
    ''' doesn't need each individual mismatch logged, since he uses a separate PowerApps
    ''' tool to manage these mappings. RunAsync tallies this count and shows it via a
    ''' single end-of-run popup instead, matching the original's own summary message.
    ''' </summary>
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

    ''' <summary>
    ''' Four-tier fallback, in the original's exact order: recipient company in Customer_QB,
    ''' then sender company (excluding "Intelligent Office" and blank) in Customer_QB, then
    ''' recipient company in SendPro_XRef, then sender company in SendPro_XRef. Falls back
    ''' to the "1" placeholder if none resolve.
    ''' </summary>
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

    ''' <summary>
    ''' REAL BUG FIXED: Account_Num is an Int32 column in SQL Server, not a string as
    ''' originally assumed - confirmed via a real InvalidCastException on reader.GetString.
    ''' Reads via GetValue()+Convert.ToString() instead of a strict GetInt32/GetString
    ''' call, so this doesn't break again if the actual column type turns out to be some
    ''' other numeric type (same defensive approach already used elsewhere in this port
    ''' after a similar surprise with PaperCut's scan_action_id column).
    ''' </summary>
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
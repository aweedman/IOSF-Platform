Imports System.Data.Odbc
Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms
Imports ClosedXML.Excel

''' <summary>
''' Direct port of Landing Page.cls: Command74_Click() ("Kube Payments to QB"). Unlike
''' Kube Invoices, this is mostly a MANUAL-REVIEW TRIAGE TOOL, not a fully-automated
''' importer - only four scenarios (Credit Card/Debit, EFT, Direct Deposit, Check) insert
''' a payment automatically. Every other recognized scenario (refunds, NSF, prepayment
''' application, credit memo application, same-day voids, "no scenario found") just shows
''' a message telling a human what to go do manually in QuickBooks - nothing is automated
''' for those, matching the original exactly.
'''
''' Confirmed with Al before porting:
'''  - Keep the popup/MsgBox behavior for manual-review scenarios, rather than logging them
'''    to Error_Log - this is a genuine, deliberate exception to this port's usual
'''    separation between job logic and UI (every other job takes explicit parameters and
'''    leaves interaction to the caller). MessageBox.Show is called directly from this
'''    module for that reason. This also means, same as the original, this job can only
'''    meaningfully be run interactively - not from a headless/scheduled context.
'''  - The original's duplicate/void checks read from "Payments_QB", but every INSERT
'''    targets "ReceivePayment_QB" - two different table names. Confirmed: "Payments" is a
'''    QODBC SP_Report used by a different, not-yet-ported process (not a real writable
'''    table), and Al suggested checking ReceivePayment instead, which the actual INSERTs
'''    already target. This port checks ReceivePayment for both - consistent with what it
'''    inserts into, and confirmed via QODBC's own table documentation that RefNumber
'''    (VARCHAR(20)) is a real, valid column there.
'''
''' Same QODBC approach as KubeInvoicesToQbJob, and for the same reasons (see that file's
''' remarks for the full story): ONE OdbcConnection for the entire run, raw literal SQL
''' text (via QodbcHelpers.SqlLiteral/OdbcDateLiteral) rather than bound parameters,
''' everything synchronous. Table names use the same no-"_QB"-suffix convention already
''' confirmed (ReceivePayment, not ReceivePayment_QB).
'''
''' REQUIRES ADDING THE ClosedXML NUGET PACKAGE (same as KubeInvoicesToQbJob - already
''' added if that job's setup is done).
'''
''' Other deviations:
'''  - No row-grouping/aggregation needed here at all (unlike Kube Invoices) - each Excel
'''    row is an independent payment, processed in one streaming pass, closely matching
'''    the original's own single-pass loop structure.
'''  - A failure on one row is logged and does NOT stop the rest, matching the original's
'''    On Error Resume Next. The returned error count reflects only genuine SQL/exception
'''    failures (matching the original's own RoutineErrors variable) - NOT the manual-
'''    review popups, which are a normal, expected outcome of running this job, not errors.
'''  - SAFETY ADDITION (not in the original, same as Kube Invoices): stops once past the
'''    sheet's actual used range rather than relying purely on a "Total" row to Exit Do.
''' </summary>
Public Module KubePaymentsToQbJob

    Private Const NonMemberListId As String = "800004B8-1704221508"
    Private Const ArAccountRefListId As String = "8000004E-1474429127"
    Private Const DepositToAccountRefListId As String = "80000098-1475524328"
    Private Const PaymentMethod_CreditDebitCard As String = "80000010-1496788905"
    Private Const PaymentMethod_Eft As String = "8000000E-1475535789"
    Private Const PaymentMethod_Cash As String = "80000001-1472090934"
    Private Const PaymentMethod_Check As String = "80000002-1472090934"

    ''' <summary>
    ''' onCreated, if given, is invoked once per successfully created payment -
    ''' "Payment <#> for <customer>" - so the caller (LandingPageForm) can surface it in
    ''' the UI log as it happens. Optional and defaults to Nothing so headless/scheduled
    ''' dispatch call sites don't need to change. Runs on a background thread - onCreated
    ''' must NOT touch UI controls directly; the caller is responsible for marshaling back
    ''' to the UI thread.
    ''' </summary>
    Public Async Function RunAsync(excelFilePath As String, Optional onCreated As Action(Of String) = Nothing) As Task(Of Integer)
        Dim errorCount = 0
        Dim listIdByKubeAccount = Await ResolveListIdsAsync()
        Dim windowStart = New Date(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(-1)

        Using conn As New OdbcConnection(ConfigHelper.QodbcConnectionString)
            conn.Open()

            Dim existingRefs = FetchExistingRefNumbers(conn, windowStart)

            Using workbook = New XLWorkbook(excelFilePath)
                Dim ws = workbook.Worksheet("Report1")
                Dim lastRow = ws.LastRowUsed().RowNumber()
                Dim headerFound = False
                Dim row = 1

                Do While row <= lastRow
                    Dim docSeqNo = ws.Cell(row, 1).GetString()

                    If Not headerFound AndAlso docSeqNo = "Doc. Seq. No." Then
                        headerFound = True
                        row += 1
                        Continue Do
                    End If

                    If Not headerFound Then
                        row += 1
                        Continue Do
                    End If

                    If docSeqNo = "Total" Then Exit Do

                    Try
                        ProcessRow(conn, ws, row, listIdByKubeAccount, existingRefs, windowStart, onCreated)
                    Catch ex As Exception
                        ErrorLogHelper.LogError("Kube Payments to QB", $"SQL error in row {row}: {ex.Message}")
                        errorCount += 1
                    End Try

                    row += 1
                Loop
            End Using
        End Using

        Return errorCount
    End Function

    ''' <summary>
    ''' PERFORMANCE FIX per Al: the duplicate check (fires on every row) was doing a live
    ''' QODBC query per row, which is slow - confirmed by Al that QODBC's lookup-by-
    ''' RefNumber is not efficient one at a time. Bulk-fetches existing RefNumbers within
    ''' the current and previous month (per Al's suggested window, based on TxnDate) into
    ''' an in-memory HashSet once, so the common-case per-row check becomes a plain
    ''' in-memory lookup instead of a QODBC round-trip.
    '''
    ''' IMPORTANT: this set is only trustworthy for duplicate-checking a ROW whose OWN
    ''' TxnDate falls within this same window - see ProcessRow, which falls back to a live
    ''' query for any row outside it (confirmed with Al this matters: a re-processed or
    ''' older file could contain a payment whose Control value already exists in QuickBooks
    ''' as an OLDER record never loaded into this set - without the fallback, that would be
    ''' silently missed as a duplicate and re-inserted).
    ''' </summary>
    Private Function FetchExistingRefNumbers(conn As OdbcConnection, windowStart As Date) As HashSet(Of String)
        Dim sql = $"SELECT RefNumber FROM ReceivePayment WHERE TxnDate >= {QodbcHelpers.OdbcDateLiteral(windowStart)}"

        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Using cmd As New OdbcCommand(sql, conn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    If Not reader.IsDBNull(0) Then result.Add(reader.GetString(0))
                End While
            End Using
        End Using
        Return result
    End Function

    ''' <summary>
    ''' All the per-row scenario logic, in the exact same order/precedence as the original
    ''' (each scenario is mutually exclusive - the first match wins and the row is done).
    ''' </summary>
    Private Sub ProcessRow(conn As OdbcConnection, ws As IXLWorksheet, row As Integer, listIdByKubeAccount As Dictionary(Of String, String), existingRefs As HashSet(Of String), windowStart As Date, onCreated As Action(Of String))
        Dim payDate = ws.Cell(row, 3).GetDateTime()
        Dim kubeCustRaw = ws.Cell(row, 4).GetString()
        Dim kubeCust = QodbcHelpers.ParseLeadingNumeric(If(kubeCustRaw.Length > 1, kubeCustRaw.Substring(1), String.Empty)).ToString("0")
        Dim custName = ws.Cell(row, 5).GetString()
        Dim control = ws.Cell(row, 7).GetString()
        Dim amount = ws.Cell(row, 11).GetDouble()
        Dim method = ws.Cell(row, 12).GetString()
        Dim notes = ws.Cell(row, 14).GetString()

        ' CC refund reversal
        If notes.Contains("Receipt reversed by Credit Card transaction") Then
            MessageBox.Show($"Refund to RCU {custName} credit memo for credit card refund", "Kube Payments to QB")
            Return
        End If

        ' Credit card NSF
        If notes.StartsWith("Reverse receipt") Then
            Dim refNumber = "R-" & ExtractTrailingNumberAfterLastSpace(notes)
            MessageBox.Show($"Create NSF and refund to RCU for {custName} {refNumber}", "Kube Payments to QB")
            Return
        End If

        ' Apply prepayments
        If notes = "Automatically generated apply prepay receipt." Then
            MessageBox.Show($"Apply {custName} prepayment", "Kube Payments to QB")
            Return
        End If

        ' Apply credit memos
        If notes = ":Prog Gen credit application" Then
            MessageBox.Show($"Apply {custName} credit memo", "Kube Payments to QB")
            Return
        End If

        ' Skip $0
        If amount = 0 Then Return

        ' Same-day void - only pops up if the referenced original payment IS found (matches
        ' the original: "If checkRef = "" Then GoTo NextLoop" skips silently otherwise)
        If notes.StartsWith(":Prog Gen Reverses receipt") Then
            Dim match = Text.RegularExpressions.Regex.Match(notes, "(\d+)$")
            Dim refNumber = "R-" & match.Value
            Dim checkRef = LookupSingleValue(conn, $"SELECT RefNumber FROM ReceivePayment WHERE RefNumber = {QodbcHelpers.SqlLiteral(refNumber)}")
            If String.IsNullOrEmpty(checkRef) Then Return
            MessageBox.Show($"Zero out and unapply Payment {refNumber} for customer {custName}", "Kube Payments to QB")
            Return
        End If

        ' Duplicate check - skips silently, matching the original. Uses the fast
        ' in-memory set for the common case (this row's own date falls within the
        ' bulk-fetched window), but falls back to a live QODBC query for any row outside
        ' it - the in-memory set was never populated with older records, so trusting it
        ' for an out-of-window row could silently miss a real duplicate and re-insert an
        ' already-existing payment. Confirmed with Al this matters for reprocessed/older
        ' files, not just a theoretical edge case.
        Dim isDuplicate As Boolean
        If payDate >= windowStart Then
            isDuplicate = existingRefs.Contains(control)
        Else
            Dim liveCheckRef = LookupSingleValue(conn, $"SELECT RefNumber FROM ReceivePayment WHERE RefNumber = {QodbcHelpers.SqlLiteral(control)}")
            isDuplicate = Not String.IsNullOrEmpty(liveCheckRef)
        End If
        If isDuplicate Then Return

        ' QB ID resolution
        Dim listId As String = Nothing
        Dim isNonMember As Boolean
        listIdByKubeAccount.TryGetValue(kubeCust, listId)
        If String.IsNullOrEmpty(listId) Then
            listId = NonMemberListId
            isNonMember = True
        End If
        Dim autoApplyStr = If(isNonMember, "FALSE", "TRUE")

        ' CC & Debit Card
        If notes.StartsWith("Credit Card On-Line Payment") OrElse notes.StartsWith("Debit Card On-Line Payment") Then
            InsertPayment(conn, listId, control, payDate, PaymentMethod_CreditDebitCard, amount, autoApplyStr, existingRefs, custName, onCreated)
            If isNonMember Then MessageBox.Show($"Nonmember Apply {control}", "Kube Payments to QB")
            Return
        End If

        ' EFT
        If notes.StartsWith("Online Payment - EFT") Then
            InsertPayment(conn, listId, control, payDate, PaymentMethod_Eft, amount, autoApplyStr, existingRefs, custName, onCreated)
            If isNonMember Then MessageBox.Show($"Nonmember Apply {control}", "Kube Payments to QB")
            Return
        End If

        ' Direct Deposit
        If notes = "" AndAlso method = "Cash" Then
            InsertPayment(conn, listId, control, payDate, PaymentMethod_Cash, amount, autoApplyStr, existingRefs, custName, onCreated)
            If isNonMember Then MessageBox.Show($"Nonmember Apply {control}", "Kube Payments to QB")
            Return
        End If

        ' Check
        If notes = "" AndAlso method.StartsWith("Cheque:") Then
            InsertPayment(conn, listId, control, payDate, PaymentMethod_Check, amount, autoApplyStr, existingRefs, custName, onCreated)
            If isNonMember Then MessageBox.Show($"Nonmember Apply {control}", "Kube Payments to QB")
            Return
        End If

        ' NSF
        If notes.StartsWith("NSF") Then
            Dim refNumber = "R-" & ExtractTrailingNumberAfterLastSpace(notes)
            MessageBox.Show($"Create NSF for {custName} {refNumber}", "Kube Payments to QB")
            Return
        End If

        ' No scenario found
        MessageBox.Show($"No Scenario Found {control} {notes}", "Kube Payments to QB")
    End Sub

    ''' <summary>Mid(Notes, InStrRev(Notes," ")+1) then Val() - the original's exact "trailing number after the last space" extraction, used for both NSF scenarios.</summary>
    Private Function ExtractTrailingNumberAfterLastSpace(notes As String) As String
        Dim lastSpace = notes.LastIndexOf(" "c)
        Dim afterSpace = If(lastSpace >= 0, notes.Substring(lastSpace + 1), notes)
        Return QodbcHelpers.ParseLeadingNumeric(afterSpace).ToString("0")
    End Function

    Private Function LookupSingleValue(conn As OdbcConnection, sql As String) As String
        Using cmd As New OdbcCommand(sql, conn)
            Using reader = cmd.ExecuteReader()
                If reader.Read() AndAlso Not reader.IsDBNull(0) Then
                    Return reader.GetString(0)
                End If
            End Using
        End Using
        Return String.Empty
    End Function

    Private Sub InsertPayment(conn As OdbcConnection, listId As String, control As String, payDate As Date,
                               paymentMethodListId As String, amount As Double, autoApplyStr As String,
                               existingRefs As HashSet(Of String), custName As String, onCreated As Action(Of String))
        Dim sql = "INSERT INTO ReceivePayment (CustomerRefListID, ARAccountRefListID, TxnDate, RefNumber, PaymentMethodRefListID, DepositToAccountRefListID, TotalAmount, IsAutoApply) VALUES (" &
            QodbcHelpers.SqlLiteral(listId) & ", " &
            QodbcHelpers.SqlLiteral(ArAccountRefListId) & ", " &
            QodbcHelpers.OdbcDateLiteral(payDate) & ", " &
            QodbcHelpers.SqlLiteral(control) & ", " &
            QodbcHelpers.SqlLiteral(paymentMethodListId) & ", " &
            QodbcHelpers.SqlLiteral(DepositToAccountRefListId) & ", " &
            amount.ToString(Globalization.CultureInfo.InvariantCulture) & ", " &
            autoApplyStr & ")"

        Using cmd As New OdbcCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End Using

        onCreated?.Invoke($"Payment {control} for {custName}")

        ' Keep the pre-fetched set in sync with this run's own inserts - otherwise a
        ' duplicate Control value later in the SAME file wouldn't be caught, since the
        ' set was only populated once at the start (unlike the original's live per-row
        ' query, which always reflected QuickBooks' current state including this run's
        ' own earlier inserts).
        existingRefs.Add(control)
    End Sub

    ''' <summary>
    ''' SQL-Server-side only, same approach as KubeInvoicesToQbJob's identical resolution -
    ''' Evo_Customer_XRef joined to SQL Server's own Customer_QB.
    ''' </summary>
    Private Async Function ResolveListIdsAsync() As Task(Of Dictionary(Of String, String))
        Dim result As New Dictionary(Of String, String)
        Const sql As String =
            "SELECT x.KubeAccountId, c.ListID " &
            "FROM Evo_Customer_XRef x " &
            "LEFT JOIN Customer_QB c ON x.ThirdPartyAccountId = c.AccountNumber"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                Await conn.OpenAsync()
                Using reader = Await cmd.ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        If reader.IsDBNull(0) Then Continue While
                        Dim kubeAccountId = reader.GetString(0)
                        Dim listId = If(reader.IsDBNull(1), Nothing, reader.GetString(1))
                        result(kubeAccountId) = listId
                    End While
                End Using
            End Using
        End Using

        Return result
    End Function

End Module
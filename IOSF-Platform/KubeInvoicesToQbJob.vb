Imports System.Data.Odbc
Imports Microsoft.Data.SqlClient
Imports ClosedXML.Excel

''' <summary>
''' Direct port of Landing Page.cls: Command73_Click() ("Kube Invoices to QB"). Reads a
''' Kube-exported Excel file, creates new QuickBooks Invoices (for invoice groups with a
''' positive total) and Credit Memos (for groups with a zero-or-negative total), skipping
''' anything already imported. Confirmed with Al before porting: the "TxnDate >= 2026-07-01"
''' cutoff is a deliberate permanent constant (the date of the Kube cutover), not a
''' one-time testing value.
'''
''' REQUIRES ADDING THE ClosedXML NUGET PACKAGE to the project (same caveat as Ical.Net
''' elsewhere in this port: could not be compile-tested here, NuGet isn't reachable from
''' this sandbox - if Workbook/Worksheet/Cell usage below doesn't compile as written,
''' that's the first thing to check against your installed version).
'''
''' TABLE NAMES: confirmed with Al that QODBC's real table names don't carry the "_QB"
''' suffix Access used for its linked-table aliases (Item, Invoice, InvoiceLine,
''' CreditMemo, CreditMemoLine) - same pattern already established for Customer_QB's real
''' name (just "Customer") elsewhere in this port.
'''
''' ARCHITECTURAL DIFFERENCE FROM THE ORIGINAL (confirmed with Al): the original ran one
''' Access query joining QODBC-linked QuickBooks tables directly to SQL-Server-linked
''' tables (Evo_Customer_XRef_SQL, Customer_Sync_From_QB_SQL) - Access could do this
''' transparently since both were just linked tables to it. ADO.NET can't join across two
''' separate database connections in one query. This port fetches the QuickBooks-side data
''' (via QODBC) and the SQL-Server-side ListID resolution (via Evo_Customer_XRef joined to
''' SQL Server's own Customer_QB - confirmed as the real, already-established table for
''' this in CustomerMasterDeltaJob.vb) as two separate queries, then matches them in
''' memory.
'''
''' QODBC ACCESS PATTERN REWRITTEN TO LITERALLY DUPLICATE ACCESS'S OWN APPROACH, per Al,
''' after repeated real errors trying to be more ADO.NET-idiomatic:
'''  - ONE OdbcConnection for the ENTIRE job run, opened once and reused for every QODBC
'''    operation - matching how Access/DAO keeps one persistent connection open for the
'''    whole session, rather than this file's earlier versions opening a fresh connection
'''    per statement (confirmed via a real "Error parsing complete XML return string"
'''    error) or even per invoice group (same error persisted). QODBC's FQSaveToCache
'''    mechanism (confirmed by Al: line items are cached via FQSaveToCache=TRUE, then the
'''    Invoice/CreditMemo insert triggers QuickBooks to actually create the transaction)
'''    appears tied to the connection/session, and needs that session to be the SAME one
'''    across the whole run, not just within one invoice group.
'''  - RAW, LITERAL SQL TEXT for every QODBC statement, built via string concatenation with
'''    proper quote escaping - matching Access's own DoCmd.RunSQL exactly, NOT bound "?"
'''    parameters. Confirmed via a real error ("Column not found: @RefNumber") that named
'''    parameters don't work in INSERT...VALUES against QODBC at all; switching to
'''    positional "?" parameters cleared that specific error but the XML-parsing error
'''    persisted regardless of connection scoping, so parameters themselves (bound values
'''    of any kind) may interact differently with QODBC's SQL-to-qbXML translation than
'''    literal SQL text does. Per Al's direct request, this now duplicates Access's actual
'''    approach rather than a .NET-idiomatic approximation of it.
'''  - Given everything now runs sequentially against one shared, synchronous ODBC
'''    connection, the QODBC-touching helpers below are synchronous methods rather than
'''    Task.Run-wrapped - there's no real parallelism to gain from wrapping strictly
'''    sequential operations on a single connection, and it keeps the "one shared
'''    connection" invariant simpler to see directly in the code.
'''
''' OTHER DEVIATIONS (confirmed with Al):
'''  - Row grouping no longer assumes same-invoice rows arrive consecutively (the original
'''    had no ORDER BY and relied on this by accident) - done via an in-memory GroupBy
'''    instead, which doesn't depend on row order at all.
'''  - Missing charge codes are ALL reported before aborting, not just the first one (the
'''    original's loop could never reach a second one, since its MsgBox branch always hit
'''    Exit Sub on the very first iteration).
'''  - Per-row/per-group failures increment an error count and continue to the next group,
'''    matching the original's On Error Resume Next behavior - a single bad invoice group
'''    doesn't abort the whole run.
'''  - No artificial transaction wrapping around the QODBC inserts - the original didn't
'''    use one either, and QODBC (a live bridge into QuickBooks itself, not a true
'''    relational database) doesn't meaningfully support one; adding a fake transaction
'''    wrapper here would imply atomicity that doesn't actually exist against QuickBooks.
'''  - Line items are inserted BEFORE the header for a given group, not after - confirmed
'''    both by re-reading the original's actual execution order (it also does this) and by
'''    an explicit QODBC error requiring it ("insert Child/Detail record(s) before
'''    inserting Parent/Header record").
'''
''' Everything else - the exact invoice-number floor/padding algorithm, the two-format
''' invoice-number handling (12-char padded for the floor comparison, "cleaned" for the
''' actually-stored RefNumber), the "last row wins" ListID resolution per line-item group,
''' the non-member fallback account, the per-charge-code line item structure - is preserved
''' exactly as the original had it; these all look like deliberate business logic, not
''' incidental bugs, and there's no evidence otherwise.
''' </summary>
Public Module KubeInvoicesToQbJob

    Private Const NonMemberListId As String = "800004B8-1704221508"
    Private Const InvoiceClass As String = "San Francisco"
    Private Const CutoverDateLiteral As String = "2026-07-01" ' confirmed with Al: permanent, the date of the Kube cutover - not a testing value

    Private Class TempRow
        Public Property InvoiceNumber As String ' raw, as read from Excel (e.g. "202607000001")
        Public Property InvoiceDate As Date
        Public Property KubeAccount As String
        Public Property CustomerName As String
        Public Property ChargeCode As String
        Public Property Amount As Double
    End Class

    Public Async Function RunAsync(excelFilePath As String) As Task(Of Integer)
        Dim errorCount = 0

        ' ONE connection for the entire run - see class remarks for why.
        Using qbConn As New OdbcConnection(ConfigHelper.QodbcConnectionString)
            qbConn.Open()

            Dim targetFloor = ComputeTargetFloor(qbConn)

            Dim rows = ParseExcelRows(excelFilePath, targetFloor)
            If rows.Count = 0 Then Return 0 ' nothing new to import - not an error, matches the original silently doing nothing when no rows qualify

            Dim missingCodes = FindMissingChargeCodes(qbConn, rows)
            If missingCodes.Count > 0 Then
                For Each code In missingCodes
                    ErrorLogHelper.LogError("Kube Invoices to QB", $"Charge Code {code} Not Found")
                Next
                Return missingCodes.Count
            End If

            ' SQL Server side - unaffected by any of the QODBC connection/parameter work
            ' above, stays genuinely async against its own separate connection.
            Dim listIdByKubeAccount = Await ResolveListIdsAsync(rows)

            Dim existingInvoiceRefs = FetchExistingRefNumbers(qbConn, "Invoice")
            Dim existingCreditMemoRefs = FetchExistingRefNumbers(qbConn, "CreditMemo")

            Dim byInvoiceNumber = rows.GroupBy(Function(r) r.InvoiceNumber).ToList()
            Dim invoiceGroups = byInvoiceNumber.Where(Function(g) g.Sum(Function(r) r.Amount) > 0).ToList()
            Dim creditMemoGroups = byInvoiceNumber.Where(Function(g) g.Sum(Function(r) r.Amount) <= 0).ToList()

            errorCount += CreateHeaderAndLines(qbConn, invoiceGroups, listIdByKubeAccount, existingInvoiceRefs,
                headerTable:="Invoice", lineTable:="InvoiceLine",
                headerColumns:="RefNumber, CustomerRefListID, BillAddressAddr1, TxnDate",
                lineColumns:="InvoiceLineItemRefFullName, InvoiceLineRate, InvoiceLineClassRefFullName, FQSaveToCache",
                amountSign:=1)

            errorCount += CreateHeaderAndLines(qbConn, creditMemoGroups, listIdByKubeAccount, existingCreditMemoRefs,
                headerTable:="CreditMemo", lineTable:="CreditMemoLine",
                headerColumns:="RefNumber, CustomerRefListID, BillAddressAddr1, TxnDate",
                lineColumns:="CreditMemoLineItemRefFullName, CreditMemoLineRate, CreditMemoLineClassRefFullName, FQSaveToCache",
                amountSign:=-1)
        End Using

        Return errorCount
    End Function

    ''' <summary>
    ''' Replicates: DMax("Val(RefNumber)", "Invoice"/"CreditMemo", crit), then the
    ''' original's exact 12-char padding rule:
    '''   targetStr = Left(targetStr, 4) & String(12 - Len(targetStr), "0") & Mid(targetStr, 5)
    ''' QODBC has no equivalent to VBA's Val() pushable into SQL, so this fetches matching
    ''' RefNumber strings via a WHERE clause and computes the numeric max in memory instead.
    ''' </summary>
    Private Function ComputeTargetFloor(conn As OdbcConnection) As String
        Dim invoiceMax = MaxRefNumberValue(conn, "Invoice")
        Dim creditMemoMax = MaxRefNumberValue(conn, "CreditMemo")
        Dim target = Math.Max(invoiceMax, creditMemoMax)

        Dim targetStr = target.ToString("0")
        Dim left4 = If(targetStr.Length >= 4, targetStr.Substring(0, 4), targetStr)
        Dim padCount = Math.Max(0, 12 - targetStr.Length)
        Dim remainder = If(targetStr.Length >= 5, targetStr.Substring(4), String.Empty)
        Return left4 & New String("0"c, padCount) & remainder
    End Function

    Private Function MaxRefNumberValue(conn As OdbcConnection, table As String) As Double
        Dim currentYear = DateTime.Today.Year.ToString()
        Dim priorYear = (DateTime.Today.Year - 1).ToString()
        Dim sql = $"SELECT RefNumber FROM {table} WHERE (RefNumber LIKE '{currentYear}%' OR RefNumber LIKE '{priorYear}%') AND TxnDate >= {OdbcDateLiteral(DateTime.Parse(CutoverDateLiteral))}"

        Dim maxValue As Double = 0
        Using cmd As New OdbcCommand(sql, conn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    If reader.IsDBNull(0) Then Continue While
                    Dim value = ParseLeadingNumeric(reader.GetString(0))
                    If value > maxValue Then maxValue = value
                End While
            End Using
        End Using
        Return maxValue
    End Function

    ''' <summary>
    ''' Replicates the Excel-reading loop exactly: finds the "Doc. Seq. No." header row,
    ''' skips rows until then, stops at a "Total" row, skips rows with a blank column 7,
    ''' and only keeps rows whose invoice number exceeds targetFloor. Both strings are the
    ''' same fixed 12-character length by construction, so this string comparison gives the
    ''' same result as a numeric one - preserved exactly as the original had it rather than
    ''' converted to a numeric comparison.
    '''
    ''' invnum is deliberately declared OUTSIDE the loop and only updated when column 1 is
    ''' non-blank - this matches the original's "If Cells(i,1).Value <> "" Then invnum=..."
    ''' exactly: some rows leave column 1 blank for continuation line items belonging to
    ''' the SAME invoice as the row above, and invnum must carry forward across those rows
    ''' rather than reset.
    '''
    ''' SAFETY ADDITION (not in the original): stops once past the sheet's actual used
    ''' range, rather than relying purely on finding a "Total" row to Exit Do. The original
    ''' had no such bound and could loop indefinitely against a malformed export with no
    ''' Total row - this doesn't change any business logic, just prevents a genuine hang.
    ''' </summary>
    Private Function ParseExcelRows(excelFilePath As String, targetFloor As String) As List(Of TempRow)
        Dim result As New List(Of TempRow)

        Using workbook = New XLWorkbook(excelFilePath)
            Dim ws = workbook.Worksheet("Report1")
            Dim lastRow = ws.LastRowUsed().RowNumber()
            Dim headerFound = False
            Dim invnum As String = Nothing ' carries forward across blank-column-1 rows - see remarks above
            Dim row = 1

            Do While row <= lastRow
                Dim col1 = ws.Cell(row, 1).GetString()
                If col1 <> "" Then invnum = col1

                If Not headerFound AndAlso invnum = "Doc. Seq. No." Then
                    headerFound = True
                    row += 1
                    Continue Do
                End If

                If Not headerFound Then
                    row += 1
                    Continue Do
                End If

                If invnum = "Total" Then Exit Do

                If ws.Cell(row, 7).GetString() = "" Then
                    row += 1
                    Continue Do
                End If

                Dim chargeCode = ws.Cell(row, 10).GetString()
                Dim invDate = ws.Cell(row, 3).GetDateTime()
                Dim kubeCustRaw = ws.Cell(row, 4).GetString()
                Dim kubeCust = ParseLeadingNumeric(If(kubeCustRaw.Length > 1, kubeCustRaw.Substring(1), String.Empty)).ToString("0")
                Dim custName = ws.Cell(row, 5).GetString()
                Dim amount = ws.Cell(row, 13).GetDouble()

                If String.CompareOrdinal(invnum, targetFloor) > 0 Then
                    result.Add(New TempRow With {
                        .InvoiceNumber = invnum,
                        .InvoiceDate = invDate,
                        .KubeAccount = kubeCust,
                        .CustomerName = custName,
                        .ChargeCode = chargeCode,
                        .Amount = amount
                    })
                End If

                row += 1
            Loop
        End Using

        Return result
    End Function

    Private Function FindMissingChargeCodes(conn As OdbcConnection, rows As List(Of TempRow)) As List(Of String)
        Dim distinctCodes = rows.Select(Function(r) r.ChargeCode).Distinct().ToList()
        Dim existingItemNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Using cmd As New OdbcCommand("SELECT Name FROM Item", conn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    If Not reader.IsDBNull(0) Then existingItemNames.Add(reader.GetString(0))
                End While
            End Using
        End Using

        Return distinctCodes.Where(Function(c) Not existingItemNames.Contains(c)).ToList()
    End Function

    ''' <summary>
    ''' SQL-Server-side only: Evo_Customer_XRef joined to SQL Server's own Customer_QB
    ''' (confirmed as the real, already-established table for this in
    ''' CustomerMasterDeltaJob.vb - NOT the QODBC live table, despite the similar name).
    ''' Genuinely async against its own separate SQL Server connection - unaffected by any
    ''' of the QODBC-specific rework elsewhere in this file.
    ''' </summary>
    Private Async Function ResolveListIdsAsync(rows As List(Of TempRow)) As Task(Of Dictionary(Of String, String))
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

    ''' <summary>
    ''' Filtered to current/prior year and the Kube cutover date, same as
    ''' MaxRefNumberValue - confirmed with Al this can never exclude a real duplicate,
    ''' since a Kube export is guaranteed to never contain an invoice number older than
    ''' that window. Bounds the result set size regardless of how much invoice history
    ''' accumulates in QuickBooks over time.
    ''' </summary>
    Private Function FetchExistingRefNumbers(conn As OdbcConnection, table As String) As HashSet(Of String)
        Dim currentYear = DateTime.Today.Year.ToString()
        Dim priorYear = (DateTime.Today.Year - 1).ToString()
        Dim sql = $"SELECT RefNumber FROM {table} WHERE (RefNumber LIKE '{currentYear}%' OR RefNumber LIKE '{priorYear}%') AND TxnDate >= {OdbcDateLiteral(DateTime.Parse(CutoverDateLiteral))}"

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

    ''' <summary>Left(raw,4) & CStr(CLng(Mid(raw,5))) - the original's exact invoice-number "cleaning" rule.</summary>
    Private Function CleanInvoiceNumber(raw As String) As String
        Dim left4 = If(raw.Length >= 4, raw.Substring(0, 4), raw)
        Dim remainder = If(raw.Length >= 5, raw.Substring(4), "0")
        Dim numeric = CLng(ParseLeadingNumeric(remainder))
        Return left4 & numeric.ToString()
    End Function

    ''' <summary>
    ''' Shared logic for both Invoice and Credit Memo creation - the original's two blocks
    ''' were near-identical except table names and the line-item amount sign (credit memo
    ''' line amounts are negated). Per group (one per InvoiceNumber): skip if the cleaned
    ''' RefNumber already exists (duplicate check), otherwise insert one LINE ITEM row per
    ''' DISTINCT (ChargeCode) within the group whose own sub-total isn't exactly zero
    ''' FIRST (with FQSaveToCache=TRUE, per Al - this stages them in QODBC's cache), THEN
    ''' one header row (which triggers QODBC to actually create the QuickBooks
    ''' transaction from whatever's staged in the cache). A failure on one group is logged
    ''' and does NOT stop the rest, matching the original's On Error Resume Next.
    ''' </summary>
    Private Function CreateHeaderAndLines(conn As OdbcConnection, groups As List(Of IGrouping(Of String, TempRow)),
                                           listIdByKubeAccount As Dictionary(Of String, String),
                                           existingRefs As HashSet(Of String),
                                           headerTable As String, lineTable As String,
                                           headerColumns As String, lineColumns As String,
                                           amountSign As Integer) As Integer
        Dim errorCount = 0

        For Each group In groups
            Try
                Dim cleanedRefNumber = CleanInvoiceNumber(group.Key)
                If existingRefs.Contains(cleanedRefNumber) Then Continue For ' already imported

                ' Per-(ChargeCode) line items, excluding any whose own sub-total is exactly zero
                Dim lineGroups = group.GroupBy(Function(r) r.ChargeCode).
                    Select(Function(g) New With {
                        .ChargeCode = g.Key,
                        .Amount = g.Sum(Function(r) r.Amount),
                        .InvoiceDate = g.Last().InvoiceDate,   ' "last row wins" per group, matching the original's sequential-overwrite loop behavior
                        .CustomerName = g.Last().CustomerName,
                        .KubeAccount = g.Last().KubeAccount
                    }).
                    Where(Function(g) g.Amount <> 0).
                    ToList()

                If lineGroups.Count = 0 Then Continue For

                Dim maxDate = lineGroups.Max(Function(g) g.InvoiceDate)
                Dim lastListId As String = Nothing
                Dim addr As String = String.Empty

                For Each lg In lineGroups
                    Dim resolvedListId As String = Nothing
                    listIdByKubeAccount.TryGetValue(lg.KubeAccount, resolvedListId)

                    If Not String.IsNullOrEmpty(resolvedListId) Then
                        lastListId = resolvedListId
                    Else
                        lastListId = NonMemberListId
                        addr = lg.CustomerName.Replace("'", "''")
                    End If
                Next

                ' Line items first (cached via FQSaveToCache=TRUE), header last (triggers
                ' QuickBooks to actually create the transaction) - per Al, and matching
                ' both the original's own execution order and QODBC's explicit requirement.
                For Each lg In lineGroups
                    Dim lineSql = $"INSERT INTO {lineTable} ({lineColumns}) VALUES ({SqlLiteral($"KUBE:{lg.ChargeCode}")}, {(lg.Amount * amountSign).ToString(Globalization.CultureInfo.InvariantCulture)}, {SqlLiteral(InvoiceClass)}, TRUE)"
                    Using cmd As New OdbcCommand(lineSql, conn)
                        cmd.ExecuteNonQuery()
                    End Using
                Next

                Dim headerSql = $"INSERT INTO {headerTable} ({headerColumns}) VALUES ({SqlLiteral(cleanedRefNumber)}, {SqlLiteral(lastListId)}, {SqlLiteral(addr)}, {OdbcDateLiteral(maxDate)})"
                Using cmd As New OdbcCommand(headerSql, conn)
                    cmd.ExecuteNonQuery()
                End Using

            Catch ex As Exception
                ErrorLogHelper.LogError("Kube Invoices to QB", $"SQL error processing invoice group {group.Key}: {ex.Message}")
                errorCount += 1
            End Try
        Next

        Return errorCount
    End Function

End Module
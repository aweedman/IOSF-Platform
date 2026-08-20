Imports System.Data.Odbc

''' <summary>
''' Direct port of Landing Page.cls: Command58_Click() ("Class Checks") plus the
''' transaction-type selection it delegates to ("Option Button" form, ported as
''' ClassCheckTypeDialog).
'''
''' Finds QuickBooks transaction lines missing a Class assignment (a QB concept used for
''' departmental/location tracking - PnLToDbJob's own ProfitAndLossByClass report depends
''' on every line having one), so the bookkeeper can find and fix them before finalizing
''' books.
'''
''' All 9 tables confirmed real QODBC-linked tables via the actual tbldefs (not guessed):
''' BillExpenseLine, CheckExpenseLine, CreditCardCreditExpenseLine,
''' CreditCardChargeExpenseLine, CreditMemoLine, DepositLine, InvoiceLine,
''' JournalEntryLine, SalesReceiptLine (all from their own _QB-suffixed Access aliases).
'''
''' Uses raw literal SQL via QodbcHelpers, not parameterized queries - matches the
''' established, hard-won lesson elsewhere in this port that QODBC doesn't support named
''' parameters in WHERE clauses.
'''
''' Each of the 9 queries' column list and additional filters (beyond "TxnDate >= FromDate
''' AND ...ClassRefFullName IS NULL") preserved exactly as written in the original, not
''' homogenized - e.g. only Checks/Credit Card Credits/Credit Card Expenses/Journal
''' Entries also filter Amount <> 0; only Deposits excludes Undeposited Funds/Accounts
''' Receivable; only Invoices excludes "Subtotal" description lines; only Journal Entries
''' excludes memos starting with "Bill" (LIKE 'Bill*', QODBC's own JSV-adjacent wildcard
''' syntax rather than SQL Server's '%' - preserved as written, since QODBC is not
''' standard T-SQL).
'''
''' Same simple read-only grid display pattern as CopierCountsReportJob/
''' CallCountsReportJob (the original's own report is a plain "Datasheet": 1 style, no
''' custom print layout needed here unlike Spheremail Storage).
''' </summary>
Public Module ClassChecksJob

    Private Const BillsQuery As String =
        "SELECT VendorRefFullName, ExpenseLineSeqNo, TxnDate " &
        "FROM BillExpenseLine " &
        "WHERE TxnDate >= {0} AND ExpenseLineClassRefFullName IS NULL"

    Private Const ChecksQuery As String =
        "SELECT PayeeEntityRefFullName, RefNumber, ExpenseLineSeqNo, TxnDate, ExpenseLineMemo " &
        "FROM CheckExpenseLine " &
        "WHERE TxnDate >= {0} AND ExpenseLineClassRefFullName IS NULL AND Amount <> 0"

    Private Const CreditCardCreditsQuery As String =
        "SELECT PayeeEntityRefFullName, AccountRefFullName, ExpenseLineSeqNo, TxnDate, ExpenseLineAccountRefFullName " &
        "FROM CreditCardCreditExpenseLine " &
        "WHERE TxnDate >= {0} AND ExpenseLineClassRefFullName IS NULL AND Amount <> 0 " &
        "ORDER BY PayeeEntityRefFullName, TxnDate"

    Private Const CreditCardExpensesQuery As String =
        "SELECT PayeeEntityRefFullName, AccountRefFullName, ExpenseLineSeqNo, TxnDate, ExpenseLineAccountRefFullName " &
        "FROM CreditCardChargeExpenseLine " &
        "WHERE TxnDate >= {0} AND ExpenseLineClassRefFullName IS NULL AND Amount <> 0 " &
        "ORDER BY PayeeEntityRefFullName, TxnDate"

    Private Const CreditMemosQuery As String =
        "SELECT CustomerRefFullName, RefNumber, CreditMemoLineSeqNo, TxnDate, CreditMemoLineDesc " &
        "FROM CreditMemoLine " &
        "WHERE TxnDate >= {0} AND CreditMemoLineClassRefFullName IS NULL"

    Private Const DepositsQuery As String =
        "SELECT DepositToAccountRefFullName, DepositLineAccountRefFullName, TxnDate " &
        "FROM DepositLine " &
        "WHERE TxnDate >= {0} AND DepositLineAccountRefFullName <> 'Undeposited Funds' AND DepositLineAccountRefFullName <> 'Accounts Receivable' " &
        "AND DepositLineClassRefFullName IS NULL"

    Private Const InvoicesQuery As String =
        "SELECT CustomerRefFullName, RefNumber, InvoiceLineSeqNo, TxnDate, InvoiceLineDesc " &
        "FROM InvoiceLine " &
        "WHERE TxnDate >= {0} AND InvoiceLineDesc <> 'Subtotal' AND InvoiceLineClassRefFullName IS NULL"

    Private Const JournalEntriesQuery As String =
        "SELECT RefNumber, JournalLineSeqNo, TxnDate, JournalLineAccountRefFullName " &
        "FROM JournalEntryLine " &
        "WHERE TxnDate >= {0} AND JournalLineClassRefFullName IS NULL AND JournalLineAmount <> 0 AND JournalLineMemo NOT LIKE 'Bill*'"

    Private Const SalesReceiptsQuery As String =
        "SELECT CustomerRefFullName, RefNumber, SalesReceiptLineSeqNo, TxnDate, SalesReceiptLineDesc " &
        "FROM SalesReceiptLine " &
        "WHERE TxnDate >= {0} AND SalesReceiptLineClassRefFullName IS NULL"

    Private ReadOnly QueryTemplates As New Dictionary(Of Integer, String)
    Public ReadOnly TypeLabels As New Dictionary(Of Integer, String)

    Sub New()
        QueryTemplates.Add(1, BillsQuery)
        QueryTemplates.Add(2, ChecksQuery)
        QueryTemplates.Add(3, CreditCardCreditsQuery)
        QueryTemplates.Add(4, CreditCardExpensesQuery)
        QueryTemplates.Add(5, CreditMemosQuery)
        QueryTemplates.Add(6, DepositsQuery)
        QueryTemplates.Add(7, InvoicesQuery)
        QueryTemplates.Add(8, JournalEntriesQuery)
        QueryTemplates.Add(9, SalesReceiptsQuery)

        TypeLabels.Add(1, "Bills")
        TypeLabels.Add(2, "Checks")
        TypeLabels.Add(3, "Credit Card Credits")
        TypeLabels.Add(4, "Credit Card Expenses")
        TypeLabels.Add(5, "Credit Memos")
        TypeLabels.Add(6, "Deposits")
        TypeLabels.Add(7, "Invoices")
        TypeLabels.Add(8, "Journal Entries")
        TypeLabels.Add(9, "Sales Receipts")
    End Sub

    Public Function RunCheck(transactionType As Integer, fromDate As Date) As DataTable
        Dim sql = String.Format(QueryTemplates(transactionType), QodbcHelpers.OdbcDateLiteral(fromDate))

        Dim table As New DataTable()
        Using conn As New OdbcConnection(ConfigHelper.QodbcConnectionString)
            conn.Open()
            Using cmd As New OdbcCommand(sql, conn)
                Using reader = cmd.ExecuteReader()
                    table.Load(reader)
                End Using
            End Using
        End Using
        Return table
    End Function

End Module
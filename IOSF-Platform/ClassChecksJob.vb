Imports System.Data.Odbc

''' <summary>
''' Finds QuickBooks transaction lines missing a Class assignment (a QuickBooks concept
''' used for departmental/location tracking - the P&L-by-Class report depends on every
''' line having one), across 9 transaction types, so the bookkeeper can find and fix them
''' before finalizing books.
'''
''' Uses raw literal SQL (via QodbcHelpers) rather than parameterized queries, since
''' QODBC doesn't reliably support named parameters in WHERE clauses.
'''
''' Each of the 9 queries' column list and additional filters (beyond "TxnDate >= FromDate
''' AND ...ClassRefFullName IS NULL") are intentionally not identical - e.g. only Checks/
''' Credit Card Credits/Credit Card Expenses/Journal Entries also filter Amount <> 0; only
''' Deposits excludes Undeposited Funds/Accounts Receivable; only Invoices excludes
''' "Subtotal" description lines; only Journal Entries excludes memos starting with "Bill"
''' (using QODBC's own wildcard syntax, LIKE 'Bill*', not SQL Server's '%').
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
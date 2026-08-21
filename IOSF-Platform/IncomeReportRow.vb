''' <summary>
''' One row from QuickBooks' "Income by Customer Detail" report, pulled via QODBC's report
''' mechanism (see IncomeDbJob). Column types are a best guess based on typical QODBC
''' report conventions, not verified against real driver output - worth double-checking
''' once this runs against live data, particularly Amount's precision.
''' </summary>
Public Class IncomeReportRow
    Public Property TxnType As String
    Public Property RefNumber As String
    Public Property TxnDate As DateTime?
    Public Property NameAccountNumber As String
    Public Property Account As String
    Public Property TxnClass As String
    Public Property Amount As Decimal
End Class
''' <summary>
''' One row from QuickBooks' "IncomeByCustomerDetail" report, pulled via QODBC's sp_report
''' mechanism (see remarks on IncomeDbJob). Column types are a best guess based on typical
''' QODBC report conventions - not verified against real driver output. Worth checking
''' once this actually runs, particularly Amount's precision.
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
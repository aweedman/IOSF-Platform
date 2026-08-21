''' <summary>
''' A customer eligible for a PaperCut scan-to-email action - one row from the customer
''' operations data, filtered down to accounts whose service level includes scanning.
''' </summary>
Public Class ScanActionCustomer
    Public Property ContactName As String
    Public Property PrimaryOffice As String
    Public Property EmailAddress As String
    Public Property FrequentScans As Boolean
    Public Property AccountNumber As String
    Public Property ContNum As String
End Class
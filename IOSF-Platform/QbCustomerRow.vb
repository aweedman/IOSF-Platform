''' <summary>
''' One customer record from QuickBooks, read via QODBC. Shared between the delta sync
''' (only customers changed since the last run) and the full sync (every customer),
''' since both need the identical field set.
''' </summary>
Public Class QbCustomerRow
    Public Property ListId As String
    Public Property TimeCreated As DateTime?
    Public Property TimeModified As DateTime?
    Public Property Name As String
    Public Property FullName As String
    Public Property IsActive As Boolean
    Public Property CompanyName As String
    Public Property FirstName As String
    Public Property ShipAddressAddr1 As String
    Public Property ShipAddressCity As String
    Public Property ShipAddressState As String
    Public Property Email As String
    Public Property CustomerTypeRefFullName As String
    Public Property AccountNumber As String
    Public Property PreferredPaymentMethodRefFullName As String
    Public Property CustomFieldTerminationInProcess As String
End Class
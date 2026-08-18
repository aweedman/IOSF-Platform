''' <summary>
''' Replaces Spheremail_Storage_Temp, which turns out to be a genuine LOCAL Access table
''' (confirmed via its tbldefs export - no linked-table .json descriptor, unlike the
''' _SQL-suffixed tables). Its only purpose was staging rows for the bound report between
''' "compute them" and "print them" - Access reports need a bound recordsource, .NET
''' doesn't have that constraint, so this is just a plain in-memory list now. No DB
''' round-trip needed at all.
''' </summary>
Public Class SphereMailStorageRow
    Public Property MailNumber As String
    Public Property Location As String
    Public Property Customer As String
    Public Property CreatedAt As Date
    Public Property Sender As String
    Public Property Quantity As String
    Public Property PrivateMailboxNumber As String
End Class
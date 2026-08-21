''' <summary>
''' One mail item that has been sitting in long-term storage - used by both the Spheremail
''' Storage report and the Spheremail Storage email job, which share the same underlying
''' data and just present it differently (on-screen/printed report vs. email digest).
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
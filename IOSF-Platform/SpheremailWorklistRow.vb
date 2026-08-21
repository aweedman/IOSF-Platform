''' <summary>
''' One task on the Spheremail Worklist (a forward, scan, shred, etc. waiting on staff
''' action). Grouped by customer and rendered as a PDF - see ReportGenerator.
''' GenerateSphereMailWorklistPdfAsync.
''' </summary>
Public Class SpheremailWorklistRow
    Public Property MailNumber As String
    Public Property AccountNumber As String
    Public Property CustomerName As String
    Public Property ReceivedAt As Date
    Public Property Sender As String
    Public Property Quantity As String
    Public Property Task As String
    Public Property Address As String
End Class
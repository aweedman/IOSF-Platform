''' <summary>
''' Typed row for the Spheremail Worklist, replacing the earlier DataTable-based
''' approach - needed for the grouped-by-customer PDF layout (ReportGenerator.
''' GenerateSphereMailWorklistPdfAsync), which needs to sort/group in a way a plain
''' DataTable made awkward.
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
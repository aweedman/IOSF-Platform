''' <summary>
''' io.hostedsuite.com JSON is PascalCase (same family as ListReservationsResponse in
''' HostedSuiteModels.vb) - case-insensitive binding handles it, no JsonPropertyName needed.
'''
''' NOTE: this endpoint returns a raw JSON array at the root, not an object wrapping a
''' "Data" property - deserialize directly as List(Of EvoClientItem). The original's
''' "For Each Item In Response.Data" iterates VBA-Web's Response.Data directly (the whole
''' parsed body), not a nested field - my first attempt at this wrongly assumed an object
''' wrapper, which fails immediately on the root token. See CustomerXrefJob.
''' </summary>
Public Class EvoCustomField
    Public Property Value As String
End Class

Public Class EvoClientItem
    Public Property Id As String
    Public Property Name As String
    Public Property ThirdPartyAccountId As String
    Public Property CustomFields As List(Of EvoCustomField)
End Class
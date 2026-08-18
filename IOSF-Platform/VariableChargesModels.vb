Imports System.Text.Json.Serialization

''' <summary>
''' Models for HostedSuite's newer /api/charges endpoint (io2.hostedsuite.com), confirmed
''' against its own metadata endpoint (io2.hostedsuite.com/api/json/metadata?op=ListCharges)
''' before building this. Same ListResponse-with-Items/TotalPages/TotalCount pagination
''' shape already used for reception-calls in CallCountsJob.
'''
''' Only the fields actually used by VariableChargesToDbJob are mapped - the metadata page
''' lists many additional audit fields (DateCreated, DateLastModified, ArchivedById, etc.)
''' not needed here. ApiClient's default camelCase naming policy should map most of these
''' automatically, but every field is given an explicit JsonPropertyName anyway to avoid
''' any ambiguity, matching the more defensive approach used elsewhere in this port after
''' getting bitten by naming-convention surprises on other APIs (SphereMail's snake_case).
''' </summary>
Public Class ChargeInfo
    <JsonPropertyName("id")>
    Public Property Id As String
    <JsonPropertyName("clientId")>
    Public Property ClientId As String
    <JsonPropertyName("clientName")>
    Public Property ClientName As String
    <JsonPropertyName("dateOfCharge")>
    Public Property DateOfCharge As String
    <JsonPropertyName("serviceName")>
    Public Property ServiceName As String
    <JsonPropertyName("quantity")>
    Public Property Quantity As Double?
    <JsonPropertyName("cost")>
    Public Property Cost As Double?
    <JsonPropertyName("description")>
    Public Property Description As String
End Class

Public Class ChargesListResponse
    <JsonPropertyName("items")>
    Public Property Items As List(Of ChargeInfo)
    <JsonPropertyName("totalCount")>
    Public Property TotalCount As Integer
    <JsonPropertyName("totalPages")>
    Public Property TotalPages As Integer
End Class
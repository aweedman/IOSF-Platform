Imports System.Text.Json.Serialization

''' <summary>
''' Models for HostedSuite's charges list API. Same paginated list shape (Items/
''' TotalPages/TotalCount) used elsewhere for other HostedSuite endpoints.
'''
''' Only the fields VariableChargesToDbJob actually uses are mapped here - the API
''' returns several additional audit fields (creation/modification timestamps, archive
''' info, etc.) that aren't needed. Every field has an explicit JsonPropertyName even
''' though the default camelCase naming policy would map most of them automatically,
''' to avoid any ambiguity if a field name ever doesn't match the usual convention.
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
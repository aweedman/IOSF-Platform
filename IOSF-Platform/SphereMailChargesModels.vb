Imports System.Text.Json
Imports System.Text.Json.Serialization

''' <summary>
''' Models for SphereMail's GET /admin/reports/charges/detail endpoint. Uses the same
''' JsonElement-backed flexible-string pattern as SphereMailModels.vb for pmb_number
''' (rather than a plain String property) - that same vendor's API has already shown
''' inconsistent string-vs-number typing on an analogous "pmb" field elsewhere (see
''' SphereMailModels.vb's own remarks), so the same defensive handling is applied here on
''' the same grounds, not because this specific field has been confirmed to have the issue.
''' </summary>
Public Class SphereMailChargeItem
    <JsonPropertyName("date")>
    Public Property DateRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property [Date] As String
        Get
            Return JsonHelpers.ElementToString(DateRaw)
        End Get
    End Property
    <JsonPropertyName("description")>
    Public Property Description As String
End Class

Public Class SphereMailChargeGroup
    <JsonPropertyName("pmb_number")>
    Public Property PmbNumberRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property PmbNumber As String
        Get
            Return JsonHelpers.ElementToString(PmbNumberRaw)
        End Get
    End Property
    <JsonPropertyName("items")>
    Public Property Items As List(Of SphereMailChargeItem)
End Class

Public Class SphereMailChargesResponse
    <JsonPropertyName("charges")>
    Public Property Charges As List(Of SphereMailChargeGroup)
End Class
Imports System.Text.Json
Imports System.Text.Json.Serialization

''' <summary>
''' SphereMail's API returns snake_case JSON keys. ApiClient's default camelCase naming
''' policy does NOT bridge snake_case -> PascalCase automatically (that's not just a case
''' difference), so every field needs an explicit JsonPropertyName here.
'''
''' EVERY field below is a string-or-number field (see JsonHelpers.vb remarks) - each has
''' a JsonIgnore'd raw JsonElement property that actually receives the deserialized value,
''' plus a computed public String property with the original name that every caller in
''' this codebase already expects. This avoids VB.NET's inability to implement a custom
''' JsonConverter(Of String) via Utf8JsonReader (see JsonHelpers.vb).
'''
''' Originally only Pmb/Quantity/AccountNumber were converted this way, after Pmb broke
''' first. MailNumber then broke the SAME way on the next run - proving this isn't a
''' couple of isolated fields, it's a systemic inconsistency in how this API types its
''' fields. Rather than keep discovering this one field at a time, EVERY field here now
''' goes through the same flexible conversion - genuinely-textual fields (FullName, Email,
''' Sender, RefreshToken) cost nothing extra to convert this way even if they never
''' actually arrive as a JSON number, and it closes off any further surprises from this
''' API for fields not yet exercised by a test run.
''' </summary>
Public Class SphereMailAuthResponse
    <JsonPropertyName("refresh_token")>
    Public Property RefreshTokenRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property RefreshToken As String
        Get
            Return JsonHelpers.ElementToString(RefreshTokenRaw)
        End Get
    End Property
End Class

Public Class SphereMailCustomer
    <JsonPropertyName("pmb")>
    Public Property PmbRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property Pmb As String
        Get
            Return JsonHelpers.ElementToString(PmbRaw)
        End Get
    End Property
    <JsonPropertyName("full_name")>
    Public Property FullNameRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property FullName As String
        Get
            Return JsonHelpers.ElementToString(FullNameRaw)
        End Get
    End Property
    <JsonPropertyName("email")>
    Public Property EmailRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property Email As String
        Get
            Return JsonHelpers.ElementToString(EmailRaw)
        End Get
    End Property
End Class

Public Class SphereMailCustomersResponse
    <JsonPropertyName("customers")>
    Public Property Customers As List(Of SphereMailCustomer)
End Class

Public Class SphereMailMailItem
    <JsonPropertyName("mail_number")>
    Public Property MailNumberRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property MailNumber As String
        Get
            Return JsonHelpers.ElementToString(MailNumberRaw)
        End Get
    End Property
    <JsonPropertyName("received_at")>
    Public Property ReceivedAtRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property ReceivedAt As String
        Get
            Return JsonHelpers.ElementToString(ReceivedAtRaw)
        End Get
    End Property
    <JsonPropertyName("sender")>
    Public Property SenderRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property Sender As String
        Get
            Return JsonHelpers.ElementToString(SenderRaw)
        End Get
    End Property
    <JsonPropertyName("quantity")>
    Public Property QuantityRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property Quantity As String
        Get
            Return JsonHelpers.ElementToString(QuantityRaw)
        End Get
    End Property
    <JsonPropertyName("account_number")>
    Public Property AccountNumberRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property AccountNumber As String
        Get
            Return JsonHelpers.ElementToString(AccountNumberRaw)
        End Get
    End Property

    ' Fields below added for SpheremailWorklistJob's use of this same endpoint - not
    ' needed by SphereMailStorageJob, but this is the same /mail_items response shape,
    ' so extending the one shared model rather than duplicating the whole class.
    <JsonPropertyName("delivery_days")>
    Public Property DeliveryDaysRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property DeliveryDays As String
        Get
            Return JsonHelpers.ElementToString(DeliveryDaysRaw)
        End Get
    End Property

    <JsonPropertyName("account_id")>
    Public Property AccountIdRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property AccountId As String
        Get
            Return JsonHelpers.ElementToString(AccountIdRaw)
        End Get
    End Property

    <JsonPropertyName("forward_address_id")>
    Public Property ForwardAddressIdRaw As JsonElement
    <JsonIgnore>
    Public ReadOnly Property ForwardAddressId As String
        Get
            Return JsonHelpers.ElementToString(ForwardAddressIdRaw)
        End Get
    End Property
End Class

Public Class SphereMailMailItemsResponse
    <JsonPropertyName("mail_items")>
    Public Property MailItems As List(Of SphereMailMailItem)
End Class
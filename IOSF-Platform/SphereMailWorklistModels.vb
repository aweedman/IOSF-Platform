Imports System.Text.Json.Serialization

''' <summary>
''' Model for SphereMail's GET /admin/customers/{id}/forward_addresses/{id} endpoint -
''' the only genuinely new model needed for SpheremailWorklistJob. SphereMailMailItem/
''' SphereMailMailItemsResponse (for the /admin/mail_items endpoint this job also calls)
''' already existed in SphereMailModels.vb from SphereMailStorageJob's earlier use of the
''' same endpoint - extended with the three additional fields this job needs
''' (delivery_days, account_id, forward_address_id) rather than duplicated here, which
''' would have collided on class names (confirmed via a real build error).
''' </summary>
Public Class SphereMailForwardAddress
    <JsonPropertyName("street")>
    Public Property Street As String
End Class

Public Class SphereMailForwardAddressResponse
    <JsonPropertyName("forward_address")>
    Public Property ForwardAddress As SphereMailForwardAddress
End Class
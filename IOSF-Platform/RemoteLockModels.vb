Imports System.Text.Json.Serialization

''' <summary>
''' RemoteLock's API returns snake_case JSON, same as SphereMail - explicit
''' JsonPropertyName needed on every response DTO property.
''' </summary>
Public Class RemoteLockAuthResponse
    <JsonPropertyName("refresh_token")>
    Public Property RefreshToken As String
    <JsonPropertyName("access_token")>
    Public Property AccessToken As String
End Class

Public Class RemoteLockPersonAttributes
    <JsonPropertyName("name")>
    Public Property Name As String
    <JsonPropertyName("pin")>
    Public Property Pin As String
    <JsonPropertyName("department")>
    Public Property Department As String
    <JsonPropertyName("status")>
    Public Property Status As String
End Class

Public Class RemoteLockPersonItem
    <JsonPropertyName("id")>
    Public Property Id As String
    <JsonPropertyName("type")>
    Public Property Type As String
    <JsonPropertyName("attributes")>
    Public Property Attributes As RemoteLockPersonAttributes
End Class

Public Class RemoteLockMeta
    <JsonPropertyName("total_pages")>
    Public Property TotalPages As Integer
End Class

Public Class RemoteLockPersonListResponse
    <JsonPropertyName("meta")>
    Public Property Meta As RemoteLockMeta
    <JsonPropertyName("data")>
    Public Property Data As List(Of RemoteLockPersonItem)
End Class

Public Class RemoteLockPersonResponse
    <JsonPropertyName("data")>
    Public Property Data As RemoteLockPersonItem
End Class

Public Class RemoteLockAccessItem
    <JsonPropertyName("id")>
    Public Property Id As String
End Class

Public Class RemoteLockAccessListResponse
    <JsonPropertyName("data")>
    Public Property Data As List(Of RemoteLockAccessItem)
End Class
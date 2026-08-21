Imports System.Text.Json.Serialization

''' <summary>
''' Response shape for the reception-calls API. Most fields are camelCase, but the
''' duration-related fields use a distinct "...InSeconds" suffix rather than just
''' different casing, so those need explicit JsonPropertyName mapping rather than relying
''' on the default camelCase naming policy.
'''
''' Field names below are confirmed against a real API response, not assumed: the call ID
''' field is "Id" (not "GlobalId"), and the four duration fields are TalkTimeInSeconds/
''' TransferTimeInSeconds/DurationInSeconds/HoldTimeInSeconds. Getting any of these four
''' field names wrong means the property silently deserializes to 0 instead of throwing an
''' error, so any historical Call_Counts data written before these were corrected should be
''' treated as suspect for those columns and reloaded if needed.
''' </summary>
Public Class ReceptionCallItem
    <JsonPropertyName("Id")>
    Public Property Id As String
    Public Property ClientId As String
    Public Property ClientName As String
    <JsonPropertyName("TalkTimeInSeconds")>
    Public Property TalkTime As Integer
    <JsonPropertyName("TransferTimeInSeconds")>
    Public Property TransferTime As Integer
    <JsonPropertyName("DurationInSeconds")>
    Public Property Duration As Integer
    <JsonPropertyName("HoldTimeInSeconds")>
    Public Property HoldTime As Integer
    Public Property StartTime As String
    Public Property Type As String
End Class

Public Class ReceptionCallsResponse
    Public Property TotalPages As Integer
    Public Property TotalCount As Integer
    Public Property Items As List(Of ReceptionCallItem)
End Class
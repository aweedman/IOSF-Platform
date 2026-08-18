Imports System.Text.Json.Serialization

''' <summary>
''' Response shape for io2.hostedsuite.com/api/reception-calls. This API's JSON is
''' genuinely camelCase for most fields (totalPages, clientName, etc.) - unlike
''' SphereMail/RemoteLock, which use snake_case - but the numeric duration fields use a
''' distinct "...InSeconds" suffix that ApiClient's camelCase naming policy can't bridge
''' (that's a different property name, not just a case difference), so those need
''' explicit JsonPropertyName mapping.
'''
''' FIELD NAME BUG FIXED, confirmed against a real API response pasted during testing:
'''   - "GlobalId" didn't exist in the real JSON at all - the actual field is "Id".
'''   - TalkTime/TransferTime/Duration/HoldTime were all missing their "InSeconds" suffix
'''     (real fields: TalkTimeInSeconds, TransferTimeInSeconds, DurationInSeconds,
'''     HoldTimeInSeconds). Since these are different strings entirely, not just
'''     different casing, they never bound - every one of these four fields was silently
'''     deserializing to 0 for every call record processed so far. Duration/Talk/Hold/
'''     Billable values already written to Call_Counts from earlier test runs are
'''     therefore all zero and should be treated as invalid once this run is redone.
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
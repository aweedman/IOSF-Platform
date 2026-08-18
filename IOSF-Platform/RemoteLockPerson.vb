''' <summary>
''' Replaces RemoteLock_Temp, a genuine LOCAL Access table (confirmed via its tbldefs
''' export - schema-only .sql/.xml, no linked-table .json descriptor). Its only purpose
''' was caching the full RemoteLock user list in-process for PIN-based lookups during the
''' sync - a plain in-memory list serves the same purpose with no DB round-trip.
''' </summary>
Public Class RemoteLockPerson
    Public Property Id As String
    Public Property AccessName As String
    Public Property Pin As String
    Public Property Status As String
    Public Property Department As String
End Class
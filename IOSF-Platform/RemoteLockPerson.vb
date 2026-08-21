''' <summary>
''' One person from RemoteLock's full user list, cached in memory for the duration of a
''' sync so PIN-based lookups don't need a fresh API call per lookup.
''' </summary>
Public Class RemoteLockPerson
    Public Property Id As String
    Public Property AccessName As String
    Public Property Pin As String
    Public Property Status As String
    Public Property Department As String
End Class
''' <summary>
''' Converts UTC times to Pacific local time, handling PST/PDT transitions automatically
''' via the OS/ICU timezone database rather than any external time-zone service.
''' .NET 6+ resolves the IANA id cross-platform; the Windows id is a fallback for older
''' runtimes or non-ICU-enabled Windows configurations.
''' </summary>
Public Module TimeZoneHelper

    Private ReadOnly _pacificTz As TimeZoneInfo = ResolvePacificTimeZone()

    Private Function ResolvePacificTimeZone() As TimeZoneInfo
        Try
            Return TimeZoneInfo.FindSystemTimeZoneById("America/Los_Angeles") ' IANA - works cross-platform on .NET 6+
        Catch ex As TimeZoneNotFoundException
            Return TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time") ' Windows id fallback
        End Try
    End Function

    ''' <summary>Converts a UTC DateTime to Pacific local time (handles PST/PDT automatically).</summary>
    Public Function ConvertUtcToPacific(utcTime As DateTime) As DateTime
        Dim utc = DateTime.SpecifyKind(utcTime, DateTimeKind.Utc)
        Return TimeZoneInfo.ConvertTimeFromUtc(utc, _pacificTz)
    End Function

End Module
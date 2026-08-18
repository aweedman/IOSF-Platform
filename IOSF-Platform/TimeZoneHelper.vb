''' <summary>
''' Replaces the api.ipgeolocation.io/timezone/convert call in AfterHours().
''' No external dependency, no API key, handles DST automatically via the OS/ICU
''' timezone database. .NET 6+ resolves IANA IDs cross-platform; the Windows ID is
''' kept as a fallback for older runtimes or non-ICU-enabled Windows configs.
''' </summary>
Public Module TimeZoneHelper

    Private ReadOnly _pacificTz As TimeZoneInfo = ResolvePacificTimeZone()

    Private Function ResolvePacificTimeZone() As TimeZoneInfo
        Try
            Return TimeZoneInfo.FindSystemTimeZoneById("America/Los_Angeles") ' IANA - works cross-platform on .NET 6+
        Catch ex As TimeZoneNotFoundException
            Return TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time") ' Windows ID fallback
        End Try
    End Function

    ''' <summary>
    ''' Converts a UTC DateTime to Pacific local time (handles PST/PDT automatically).
    ''' Direct replacement for the Response2.Data("converted_time") value used in AfterHours().
    ''' </summary>
    Public Function ConvertUtcToPacific(utcTime As DateTime) As DateTime
        Dim utc = DateTime.SpecifyKind(utcTime, DateTimeKind.Utc)
        Return TimeZoneInfo.ConvertTimeFromUtc(utc, _pacificTz)
    End Function

End Module
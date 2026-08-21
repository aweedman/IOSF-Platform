''' <summary>
''' Shared helpers for building QODBC queries as raw SQL text rather than parameterized
''' queries. QODBC (the QuickBooks ODBC driver) doesn't reliably support named parameters
''' in compound WHERE clauses, so affected queries build literal SQL strings instead -
''' these helpers keep that string-building consistent and properly escaped.
''' </summary>
Public Module QodbcHelpers

    ''' <summary>Wraps a string value as a quoted, escaped SQL literal (doubles any embedded single quotes).</summary>
    Public Function SqlLiteral(s As String) As String
        Return "'" & If(s, String.Empty).Replace("'", "''") & "'"
    End Function

    ''' <summary>ODBC-standard date literal escape sequence - {d 'yyyy-mm-dd'}.</summary>
    Public Function OdbcDateLiteral(d As Date) As String
        Return "{d '" & d.ToString("yyyy-MM-dd") & "'}"
    End Function

    ''' <summary>Parses a leading numeric prefix from a string, ignoring anything after it (e.g. "42abc" -> 42), returning 0 if nothing numeric is found.</summary>
    Public Function ParseLeadingNumeric(s As String) As Double
        If String.IsNullOrEmpty(s) Then Return 0
        Dim sb As New Text.StringBuilder()
        Dim i = 0
        While i < s.Length AndAlso Char.IsWhiteSpace(s(i))
            i += 1
        End While
        If i < s.Length AndAlso (s(i) = "+"c OrElse s(i) = "-"c) Then
            sb.Append(s(i))
            i += 1
        End If
        Dim seenDot = False
        While i < s.Length AndAlso (Char.IsDigit(s(i)) OrElse (s(i) = "."c AndAlso Not seenDot))
            If s(i) = "."c Then seenDot = True
            sb.Append(s(i))
            i += 1
        End While
        Dim result As Double
        Double.TryParse(sb.ToString(), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, result)
        Return result
    End Function

End Module
''' <summary>
''' Shared helpers for QODBC work that duplicates Access's own DoCmd.RunSQL approach
''' (raw literal SQL text, not parameterized) - see KubeInvoicesToQbJob's class remarks
''' for why this approach was settled on. Factored out here once a second job
''' (KubePaymentsToQbJob) needed the exact same helpers, rather than duplicating them.
''' </summary>
Public Module QodbcHelpers

    ''' <summary>Wraps a string value as a quoted, escaped SQL literal - matches Access's own Replace("'","''") escaping pattern.</summary>
    Public Function SqlLiteral(s As String) As String
        Return "'" & If(s, String.Empty).Replace("'", "''") & "'"
    End Function

    ''' <summary>ODBC-standard date literal escape sequence - {d 'yyyy-mm-dd'}.</summary>
    Public Function OdbcDateLiteral(d As Date) As String
        Return "{d '" & d.ToString("yyyy-MM-dd") & "'}"
    End Function

    ''' <summary>Mimics VBA's Val(): parses a leading numeric prefix, ignoring anything after it.</summary>
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
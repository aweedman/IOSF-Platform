Imports System.Text.Json

''' <summary>
''' Some of the external APIs this app talks to are inconsistent about whether a given
''' field comes back as a JSON string or a raw JSON number (the same field can vary by
''' response). Rather than a custom JsonConverter, affected model properties are typed as
''' JsonElement (which accepts either token type with no special handling), and this
''' module converts that JsonElement into a plain string afterward - covering both cases
''' with one small helper instead of a converter per model.
''' </summary>
Public Module JsonHelpers

    Public Function ElementToString(element As JsonElement) As String
        Select Case element.ValueKind
            Case JsonValueKind.String
                Return element.GetString()
            Case JsonValueKind.Number
                Return element.GetRawText()
            Case JsonValueKind.Null, JsonValueKind.Undefined
                Return Nothing
            Case Else
                Return element.GetRawText()
        End Select
    End Function

End Module
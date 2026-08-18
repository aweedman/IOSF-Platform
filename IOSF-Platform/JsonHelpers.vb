Imports System.Text.Json

''' <summary>
''' REVISED APPROACH: the original version of this file implemented JsonConverter(Of
''' String) with an overridden Read(ByRef reader As Utf8JsonReader, ...) - this compiles
''' fine in C# but NOT in VB.NET. Utf8JsonReader is a "ref struct" (it wraps a Span(Of
''' Byte) internally), and VB.NET's compiler does not support implementing methods that
''' take a ref struct as a parameter, even though the underlying .NET type is identical
''' across both languages - a real VB.NET-specific limitation, not a mistake in the
''' original design.
'''
''' Fixed by side-stepping Utf8JsonReader entirely: System.Text.Json can deserialize any
''' JSON token directly into a JsonElement property with NO custom converter needed at
''' all (JsonElement is a normal, non-ref struct). This module just converts an
''' already-deserialized JsonElement into a canonical String afterward, handling the case
''' where the API returns a field as a JSON string in some responses and a raw JSON number
''' in others (confirmed against a real SphereMail response: "pmb" comes back as a number).
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
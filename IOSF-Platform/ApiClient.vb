Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Text.Json

''' <summary>
''' Shared HTTP client for every external API this app calls. One HttpClient instance is
''' reused for the process lifetime, which avoids socket exhaustion under repeated calls.
'''
''' DESIGN NOTE: this does NOT throw on non-2xx responses by default. Some call sites need
''' to distinguish between specific status codes (e.g. "No Content" vs. "Created") rather
''' than treating every 2xx the same way, so callers inspect response.StatusCode
''' themselves. Use response.EnsureSuccess() at call sites that only care whether the
''' request succeeded at all.
''' </summary>
Public Module ApiClient

    Private ReadOnly _httpClient As HttpClient = CreateClient()

    Private ReadOnly _jsonOptions As New JsonSerializerOptions With {
        .PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        .PropertyNameCaseInsensitive = True
    }

    Private Function CreateClient() As HttpClient
        Dim client As New HttpClient()
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
        ' Some APIs (or a WAF/CDN in front of them) reject requests with no User-Agent
        ' header at all, which .NET's HttpClient doesn't send by default - confirmed via a
        ' 403 Forbidden (HTML body, not a JSON API error) from one provider for a request
        ' that otherwise had identical headers/credentials/payload to a working one.
        client.DefaultRequestHeaders.UserAgent.ParseAdd("IOSF-Platform/1.0")
        Return client
    End Function

    ''' <summary>Wraps an HTTP response: StatusCode for exact checks, DataAs(Of T) / DataAsDocument for the body.</summary>
    Public Class ApiResponse
        Public Property StatusCode As HttpStatusCode
        Public Property Body As String

        Public ReadOnly Property IsSuccess As Boolean
            Get
                Return CInt(StatusCode) >= 200 AndAlso CInt(StatusCode) < 300
            End Get
        End Property

        Public Function DataAs(Of T)() As T
            If String.IsNullOrEmpty(Body) Then Return Nothing
            Return JsonSerializer.Deserialize(Of T)(Body, _jsonOptions)
        End Function

        Public Function DataAsDocument() As JsonDocument
            Return JsonDocument.Parse(Body)
        End Function

        ''' <summary>Throws if the response wasn't successful (2xx) - for call sites that just want "did this work at all" rather than an exact status-code check.</summary>
        Public Function EnsureSuccess() As ApiResponse
            If Not IsSuccess Then
                Throw New HttpRequestException($"Request failed with status {CInt(StatusCode)} ({StatusCode}). Body: {Body}")
            End If
            Return Me
        End Function
    End Class

    ''' <summary>Core send used by all verbs. timeoutSeconds is set per-call rather than globally, since different endpoints warrant different timeouts.</summary>
    Public Async Function SendAsync(method As HttpMethod, url As String,
                                     Optional payload As Object = Nothing,
                                     Optional headers As IDictionary(Of String, String) = Nothing,
                                     Optional timeoutSeconds As Integer = 60) As Task(Of ApiResponse)

        Using request As New HttpRequestMessage(method, url)
            If payload IsNot Nothing Then
                Dim json = JsonSerializer.Serialize(payload, _jsonOptions)
                request.Content = New StringContent(json, Encoding.UTF8, "application/json")
            End If

            If headers IsNot Nothing Then
                For Each kvp In headers
                    request.Headers.TryAddWithoutValidation(kvp.Key, kvp.Value)
                Next
            End If

            Using cts As New Threading.CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds))
                Using httpResponse = Await _httpClient.SendAsync(request, cts.Token)
                    Dim body = Await httpResponse.Content.ReadAsStringAsync()
                    Return New ApiResponse With {
                        .StatusCode = httpResponse.StatusCode,
                        .Body = body
                    }
                End Using
            End Using
        End Using
    End Function

    ' --- Convenience wrappers ---

    Public Async Function PostAsync(url As String, payload As Object,
                                     Optional headers As IDictionary(Of String, String) = Nothing,
                                     Optional timeoutSeconds As Integer = 60) As Task(Of ApiResponse)
        Return Await SendAsync(HttpMethod.Post, url, payload, headers, timeoutSeconds)
    End Function

    Public Async Function GetAsync(url As String,
                                    Optional queryParams As IDictionary(Of String, String) = Nothing,
                                    Optional headers As IDictionary(Of String, String) = Nothing,
                                    Optional timeoutSeconds As Integer = 60) As Task(Of ApiResponse)
        Dim finalUrl = url
        If queryParams IsNot Nothing AndAlso queryParams.Count > 0 Then
            Dim qs = String.Join("&", queryParams.Select(Function(kvp) $"{Uri.EscapeDataString(kvp.Key)}={Uri.EscapeDataString(kvp.Value)}"))
            finalUrl &= If(url.Contains("?"), "&", "?") & qs
        End If
        Return Await SendAsync(HttpMethod.Get, finalUrl, Nothing, headers, timeoutSeconds)
    End Function

    Public Async Function PutAsync(url As String, payload As Object,
                                    Optional headers As IDictionary(Of String, String) = Nothing,
                                    Optional timeoutSeconds As Integer = 60) As Task(Of ApiResponse)
        Return Await SendAsync(HttpMethod.Put, url, payload, headers, timeoutSeconds)
    End Function

    Public Async Function DeleteAsync(url As String,
                                       Optional headers As IDictionary(Of String, String) = Nothing,
                                       Optional timeoutSeconds As Integer = 60) As Task(Of ApiResponse)
        Return Await SendAsync(HttpMethod.Delete, url, Nothing, headers, timeoutSeconds)
    End Function

End Module
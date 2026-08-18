Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Text.Json

''' <summary>
''' Replacement for WebClient.cls / WebRequest.cls / WebResponse.cls (the VBA-Web library).
''' One shared, reused HttpClient for the process lifetime avoids socket exhaustion.
'''
''' DESIGN NOTE: this does NOT throw on non-2xx responses, by design. The original VBA
''' code checks exact status codes explicitly (e.g. "<> WebStatusCode.NoContent",
''' "<> WebStatusCode.created") rather than treating every 2xx as success - most visibly
''' in the RemoteLock provisioning flow in Landing Page.cls. Auto-throwing on "not 2xx"
''' would paper over that distinction. Callers inspect response.StatusCode themselves,
''' same as the original Response.StatusCode checks. Use response.EnsureSuccess() only
''' where the original genuinely just wanted "did this work at all" (e.g. EarlyMeeting).
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
        ' .NET's HttpClient sends no User-Agent by default, unlike Postman (which always
        ' sends its own). Some APIs, or a WAF/CDN in front of them, reject requests with no
        ' User-Agent at all - added after api.spheremail.co returned a 403 Forbidden (HTML
        ' body, not a JSON API error - consistent with a WAF-level rejection) for a request
        ' that worked fine from Postman with identical credentials and payload.
        client.DefaultRequestHeaders.UserAgent.ParseAdd("IOSF-Platform/1.0")
        Return client
    End Function

    ''' <summary>
    ''' Wraps an HTTP response. Mirrors WebResponse: StatusCode for exact checks,
    ''' DataAs(Of T) / DataAsDocument for the body.
    ''' </summary>
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

        ''' <summary>
        ''' For call sites that just want "throw if this didn't succeed" instead of an
        ''' exact status-code check (e.g. EarlyMeeting's original "Err.Raise -1" on non-OK).
        ''' </summary>
        Public Function EnsureSuccess() As ApiResponse
            If Not IsSuccess Then
                Throw New HttpRequestException($"Request failed with status {CInt(StatusCode)} ({StatusCode}). Body: {Body}")
            End If
            Return Me
        End Function
    End Class

    ''' <summary>
    ''' Core send used by all verbs. timeoutSeconds mirrors Client.TimeoutMs, which the
    ''' original set per-call (15s/20s/60s depending on endpoint) rather than globally.
    ''' </summary>
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

    ' --- Convenience wrappers matching how the app actually calls things ---

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
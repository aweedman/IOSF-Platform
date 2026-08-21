''' <summary>
''' Simple rate limiter for API calls that need to pause periodically to stay under a
''' provider's rate limit: after every 100 calls, waits 70 seconds before continuing.
''' </summary>
Public Module ApiThrottle

    Public Async Function ThrottleIfNeededAsync(counter As Integer) As Task(Of Integer)
        counter += 1
        If counter > 100 Then
            Await Task.Delay(TimeSpan.FromSeconds(70))
            counter = 0
        End If
        Return counter
    End Function

End Module
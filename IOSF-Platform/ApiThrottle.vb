''' <summary>
''' Replaces the repeated "apicounter > 100 -> wait 70s -> reset" block that appears
''' verbatim about 9 times in the original RemoteLock Users routine. Same threshold/wait
''' as the original (DoEvents busy-loop replaced with a plain async delay - no UI thread
''' to keep responsive to worry about here the way Access's was).
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
''' <summary>
''' SphereMail authentication - repeated identically in Spheremail_Customers(),
''' Spheremail_Storage(), and Command66_Click() (SphereMail Worklist, not yet ported).
''' Pulled out to one place now rather than porting the duplication three times.
''' </summary>
Public Module SphereMailAuth

    Private Const BaseUrl As String = "https://api.spheremail.co/v1"

    Public Async Function GetTokenAsync() As Task(Of String)
        Dim password = ConfigHelper.GetConfigValue("Spheremail Pass")
        Dim payload = New With {.login = "iosfadmin", .password = password}

        Dim response = Await ApiClient.PostAsync($"{BaseUrl}/authentication", payload, timeoutSeconds:=60)
        response.EnsureSuccess() ' original: Err.Raise -1 / MsgBox "API Call Error" on failure

        Return response.DataAs(Of SphereMailAuthResponse)().RefreshToken
    End Function

End Module
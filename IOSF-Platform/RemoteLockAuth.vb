''' <summary>
''' RemoteLock's OAuth2 refresh-token exchange. Extracted since this exact flow is used
''' both by RemoteLockUsersJob and (not yet ported) Command65_Click, "RemoteLock Refresh
''' Token" - a standalone button that appears to exist just to refresh the stored token
''' without running a full user sync.
'''
''' IMPORTANT: every call to this consumes the CURRENT refresh token and issues a new one
''' (standard OAuth2 refresh-token rotation) - RemoteLock invalidates the old token once
''' the new one is issued. The caller MUST persist NewRefreshToken via
''' ConfigHelper.SetConfigValue before this is called again, or the next sync will fail
''' with an invalid/already-used refresh token. RemoteLockUsersJob does this immediately;
''' keep that ordering if this gets reused elsewhere.
''' </summary>
Public Module RemoteLockAuth

    Private Const TokenUrl As String = "https://connect.remotelock.com/oauth/token"

    Public Class TokenResult
        Public Property AccessToken As String
        Public Property NewRefreshToken As String
    End Class

    Public Async Function RefreshTokenAsync() As Task(Of TokenResult)
        Dim clientId = ConfigHelper.GetConfigValue("RemoteLock Client ID")
        Dim clientSecret = ConfigHelper.GetConfigValue("RemoteLock Client Secret")
        Dim refreshToken = ConfigHelper.GetConfigValue("RemoteLock Refresh Token")

        Dim payload = New With {
            .client_id = clientId,
            .client_secret = clientSecret,
            .refresh_token = refreshToken,
            .grant_type = "refresh_token"
        }

        Dim response = Await ApiClient.PostAsync(TokenUrl, payload, timeoutSeconds:=60)
        response.EnsureSuccess() ' original: Err.Raise -1 on non-OK

        Dim data = response.DataAs(Of RemoteLockAuthResponse)()

        ' Persist immediately - see remarks above re: refresh-token rotation.
        ConfigHelper.SetConfigValue("RemoteLock Refresh Token", data.RefreshToken)

        Return New TokenResult With {
            .AccessToken = data.AccessToken,
            .NewRefreshToken = data.RefreshToken
        }
    End Function

    ''' <summary>
    ''' Direct port of the API-calling/persistence portion of Landing Page.cls:
    ''' Command65_Click ("RemoteLock Refresh Token" button) - a DIFFERENT OAuth2 grant
    ''' than RefreshTokenAsync above. This is the one-time/occasional manual
    ''' re-authorization flow (authorization_code grant): a human walks through
    ''' RemoteLock's developer portal to get a fresh Client ID/Secret/Code, then this
    ''' exchanges that Code for the FIRST refresh token and persists all three to Config.
    '''
    ''' NOT PORTED HERE: the original's five MsgBox instructional prompts and three
    ''' InputBox prompts (collecting Client ID/Secret/Code from the person running this)
    ''' are pure UI and belong with the WinForms click handler for this button - this
    ''' function just takes the three values as parameters once the person has them.
    ''' </summary>
    Public Async Function ExchangeAuthorizationCodeAsync(clientId As String, clientSecret As String, code As String) As Task
        Dim payload = New With {
            .code = code,
            .client_id = clientId,
            .client_secret = clientSecret,
            .redirect_uri = "urn:ietf:wg:oauth:2.0:oob",
            .grant_type = "authorization_code"
        }

        Dim response = Await ApiClient.PostAsync(TokenUrl, payload, timeoutSeconds:=60)
        response.EnsureSuccess()

        Dim data = response.DataAs(Of RemoteLockAuthResponse)()

        ConfigHelper.SetConfigValue("RemoteLock Client ID", clientId)
        ConfigHelper.SetConfigValue("RemoteLock Client Secret", clientSecret)
        ConfigHelper.SetConfigValue("RemoteLock Refresh Token", data.RefreshToken)
    End Function

End Module
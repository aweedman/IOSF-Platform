''' <summary>
''' Shared across every job that calls HostedSuite's newer io2.hostedsuite.com API family
''' (Call Counts, Variable Charges, and presumably more as Al cuts additional jobs over to
''' it - see conversation this was factored out in). Computes the Basic-auth-style
''' Authorization header fresh from the "Evo Pass" Config value, rather than storing a
''' separately pre-computed header value - confirmed by decoding the header this replaced
''' for Call Counts: it was Base64("sanfran:" & EvoPassValue), and "BIg7%lY8" matched the
''' real Evo Pass value exactly. One fewer place a credential can go stale or leak
''' independently - rotating the Evo password alone is sufficient, no separate manual
''' re-encoding step into Config needed.
''' </summary>
Public Module HostedSuiteAuth

    Private Const AuthUsername As String = "sanfran" ' fixed, not a Config value

    Public Function ComputeAuthHeader() As String
        Dim evoPass = ConfigHelper.GetConfigValue("Evo Pass")
        Dim credentials = $"{AuthUsername}:{evoPass}"
        Dim encoded = Convert.ToBase64String(Text.Encoding.UTF8.GetBytes(credentials))
        Return $"IO {encoded}"
    End Function

End Module
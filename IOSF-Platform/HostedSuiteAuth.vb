''' <summary>
''' Shared across every job that calls HostedSuite's newer io2.hostedsuite.com API family
''' (Call Counts, Variable Charges, and presumably more as Al cuts additional jobs over to
''' it - see conversation this was factored out in). Computes the Basic-auth-style
''' Authorization header fresh from the "Evo Pass" Config value, rather than storing a
''' separately pre-computed header value.
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
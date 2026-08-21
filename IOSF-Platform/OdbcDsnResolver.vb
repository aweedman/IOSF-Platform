Imports Microsoft.Win32

''' <summary>
''' Reads the "Server" value directly out of a System DSN's registry configuration,
''' rather than connecting through ODBC at all. The "Intelligent Office" DSN is set up,
''' by that exact name, on every machine that runs this app, but the server address each
''' machine's own local DSN points to can differ from machine to machine. Resolving the
''' server this way means appsettings.json never needs a machine-specific value - every
''' machine's config file can be identical.
'''
''' System DSNs are stored in the registry at
''' HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\{DSN name}, with a "Server" value for
''' Microsoft's own SQL Server ODBC drivers (consistent across ODBC Driver 17/18 for SQL
''' Server).
'''
''' This assumes the process runs as 64-bit, matching the DSN's own 64-bit platform. A
''' 32-bit process on 64-bit Windows would see a different registry location (Wow6432Node)
''' due to registry redirection - not handled here, since this app runs 64-bit throughout.
''' If it ever needs to run 32-bit, that's the first thing to revisit.
''' </summary>
Public Module OdbcDsnResolver

    Public Function ReadDsnServerName(dsnName As String) As String
        Dim keyPath = $"SOFTWARE\ODBC\ODBC.INI\{dsnName}"
        Using key = Registry.LocalMachine.OpenSubKey(keyPath)
            If key Is Nothing Then
                Throw New InvalidOperationException(
                    $"System DSN '{dsnName}' not found in the registry at HKEY_LOCAL_MACHINE\{keyPath}. " &
                    "Confirm it's set up as a SYSTEM (not User) DSN on this machine, matching the name exactly.")
            End If

            Dim server = TryCast(key.GetValue("Server"), String)
            If String.IsNullOrEmpty(server) Then
                Throw New InvalidOperationException(
                    $"DSN '{dsnName}' was found in the registry, but has no 'Server' value - " &
                    "check its configuration in the ODBC Data Source Administrator.")
            End If

            Return server
        End Using
    End Function

End Module
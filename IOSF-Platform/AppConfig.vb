Imports System.IO
Imports System.Text.Json

''' <summary>
''' Loads ConfigHelper.ConnectionString / QodbcConnectionString from appsettings.json
''' sitting next to the .exe. Deliberately lightweight (plain System.Text.Json, no
''' Microsoft.Extensions.Configuration) since this app only needs two string values -
''' pulling in the full Configuration framework would be more machinery than the problem
''' needs.
'''
''' IMPORTANT: appsettings.json will contain real SQL Server / QODBC credentials once
''' filled in. Given the SQL Server password that was previously committed in plaintext
''' to this repo's db-connection.json (flagged back at the very start of this port), add
''' appsettings.json to .gitignore BEFORE putting real credentials in it. Keep
''' appsettings.template.json (no real values) in source control instead, and copy it to
''' appsettings.json locally / on each deployment target.
''' </summary>
Public Module AppConfig

    Public Sub Load()
        Dim settingsPath As String = Path.Combine(AppContext.BaseDirectory, "appsettings.json")

        If Not File.Exists(settingsPath) Then
            Throw New FileNotFoundException(
                $"appsettings.json not found at {settingsPath}. Copy appsettings.template.json to " &
                "appsettings.json next to the .exe and fill in real connection strings.")
        End If

        Dim json = File.ReadAllText(settingsPath)
        Using doc = JsonDocument.Parse(json)
            Dim connectionStrings = doc.RootElement.GetProperty("ConnectionStrings")
            ConfigHelper.ConnectionString = connectionStrings.GetProperty("SqlServer").GetString()
            ConfigHelper.QodbcConnectionString = connectionStrings.GetProperty("Qodbc").GetString()
        End Using
    End Sub

End Module
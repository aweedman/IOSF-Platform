Imports Microsoft.Data.SqlClient

''' <summary>
''' Replacement for Access's DLookup("Low", "Config", "Name = '...'") calls.
''' Config is assumed to still be a linked SQL Server table (Name, Low, ... columns).
''' Adjust the connection string source (appsettings, env var, etc.) to fit how you're
''' managing config in the .NET app.
''' </summary>
Public Module ConfigHelper

    ' TODO: wire these to your actual connection string source (do not hardcode).
    ' SQL Server (Config, Holidays, IO_Employees, RemoteLock_Events, Customer_Ops_All, etc.)
    Public Property ConnectionString As String = String.Empty
    ' QODBC DSN connection to QuickBooks (Invoice, Customer, IncomeByCustomerDetail via
    ' sp_report, and any other _QB-suffixed tables). SP_Shell's original export showed a
    ' differently-named DSN ("QuickBooks Data 64" vs. "QuickBooks Data 64-Bit QRemote"
    ' used elsewhere) - confirmed these point at the same QuickBooks connection, just
    ' named differently at different points in time, so everything uses this one property.
    Public Property QodbcConnectionString As String = String.Empty

    ''' <summary>
    ''' Returns every Config row whose Name starts with the given prefix, keyed by the
    ''' REMAINDER of the name after the prefix (e.g. prefix "Early Meeting Calendar - " on
    ''' a row named "Early Meeting Calendar - Meeting Room" yields key "Meeting Room").
    ''' Used for the room-calendar list in EarlyMeetingJob, so rooms can be added or
    ''' removed by editing Config alone - no code change needed to pick up a new room.
    ''' </summary>
    Public Function GetConfigValuesByPrefix(prefix As String) As Dictionary(Of String, String)
        Const sql As String = "SELECT Name, Low FROM Config WHERE Name LIKE @Prefix + '%'"
        Dim result As New Dictionary(Of String, String)

        Using conn As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Prefix", prefix)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim fullName = reader.GetString(0)
                        Dim label = fullName.Substring(prefix.Length)
                        Dim value = If(reader.IsDBNull(1), String.Empty, reader.GetString(1))
                        result(label) = value
                    End While
                End Using
            End Using
        End Using

        Return result
    End Function

    ''' <summary>
    ''' Returns the "Low" column value from Config for the given config Name.
    ''' Mirrors DLookup("Low", "Config", "Name = 'x'") but parameterized.
    ''' </summary>
    Public Function GetConfigValue(name As String) As String
        Const sql As String = "SELECT Low FROM Config WHERE Name = @Name"

        Using conn As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Name", name)
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                Return If(result Is Nothing OrElse result Is DBNull.Value, String.Empty, result.ToString())
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Replacement for Access's "UPDATE Config_SQL SET Low = ... WHERE Name = ...".
    ''' Used to persist the rotated RemoteLock refresh token and the sync delta watermark.
    ''' </summary>
    Public Sub SetConfigValue(name As String, value As String)
        Const sql As String = "UPDATE Config SET Low = @Value WHERE Name = @Name"

        Using conn As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Value", value)
                cmd.Parameters.AddWithValue("@Name", name)
                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Papercut-related tables (scan actions, tbl_user) live in a DIFFERENT database on
    ''' the same SQL Server instance as ConnectionString (confirmed with Al: same server,
    ''' just DATABASE=Papercut instead of DATABASE=Staging). Uses SqlConnectionStringBuilder
    ''' to override just the catalog rather than assuming a specific connection string
    ''' format to string-replace.
    ''' </summary>
    Public Function GetPapercutConnectionString() As String
        Dim builder As New SqlConnectionStringBuilder(ConnectionString) With {
            .InitialCatalog = "Papercut"
        }
        Return builder.ConnectionString
    End Function

End Module
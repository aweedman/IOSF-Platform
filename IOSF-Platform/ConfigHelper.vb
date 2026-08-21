Imports Microsoft.Data.SqlClient

''' <summary>
''' Reads and writes rows in the Config table (Name/Low columns), and holds the two
''' connection strings the rest of the app uses. ConnectionString/QodbcConnectionString
''' are populated once at startup by AppConfig.Load(); everything else in this module
''' reads or writes actual Config rows.
''' </summary>
Public Module ConfigHelper

    ' SQL Server connection (Config, Holidays, IO_Employees, Customer_Ops, etc.)
    Public Property ConnectionString As String = String.Empty
    ' QODBC connection to QuickBooks (Invoice, Customer, IncomeByCustomerDetail, and
    ' every other QuickBooks-backed table this app reads).
    Public Property QodbcConnectionString As String = String.Empty

    ''' <summary>
    ''' Returns every Config row whose Name starts with the given prefix, keyed by the
    ''' remainder of the name after the prefix (e.g. prefix "Early Meeting Calendar - " on
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

    ''' <summary>Returns the "Low" column value from Config for the given config Name, or an empty string if not found.</summary>
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

    ''' <summary>Updates the "Low" column value for a given config Name. Used to persist the rotated RemoteLock refresh token and the customer-sync watermark date, among others.</summary>
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
    ''' PaperCut's tables (scan actions, tbl_user) live in a separate database on the same
    ''' SQL Server instance as ConnectionString - this returns a connection string pointing
    ''' at that database instead, by overriding just the catalog rather than assuming a
    ''' specific connection-string format to string-replace.
    ''' </summary>
    Public Function GetPapercutConnectionString() As String
        Dim builder As New SqlConnectionStringBuilder(ConnectionString) With {
            .InitialCatalog = "Papercut"
        }
        Return builder.ConnectionString
    End Function

End Module
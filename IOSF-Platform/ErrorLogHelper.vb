Imports Microsoft.Data.SqlClient

''' <summary>
''' Shared Error_Log insert, extracted rather than repeating the same INSERT block
''' per job. Column names verified against the real table where possible: Error_Log_SQL's
''' primary key is [Time], [Process] (confirmed via its tbldefs .json) - my original guess
''' of (LogDate, Routine, Message) was wrong on the first two, caught via a real build/run
''' error. The third column (the actual message text) isn't part of the primary key, so
''' it isn't revealed by that export - "Message" below is still a guess pending
''' confirmation against the real table (SSMS: SELECT TOP 1 * FROM Error_Log, or Design).
''' </summary>
Public Module ErrorLogHelper

    Public Sub LogError(routine As String, message As String)
        ' TODO: confirm the real message column name and update below if "Message" is wrong.
        Const sql As String =
            "INSERT INTO Error_Log ([Time], [Process], Message) VALUES (@Time, @Process, @Message)"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Time", DateTime.Now)
                cmd.Parameters.AddWithValue("@Process", routine)
                cmd.Parameters.AddWithValue("@Message", message)
                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

End Module
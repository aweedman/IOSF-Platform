Imports Microsoft.Data.SqlClient

''' <summary>Writes one row to Error_Log, so every job can report a problem without repeating this insert.</summary>
Public Module ErrorLogHelper

    Public Sub LogError(routine As String, message As String)
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
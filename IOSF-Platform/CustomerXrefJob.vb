Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command17_Click() ("Evo Customer XRef to DB").
''' Pulls the client list from Evo/HostedSuite and upserts it into Evo_Customer_XRef,
''' updating the Name on existing rows and inserting new ones.
'''
''' Changes from the VBA original:
'''  - Auth/initial-fetch failure aborts the whole run rather than continuing - same
'''    policy as every other job (confirmed explicitly for RemoteLock; applying the same
'''    reasoning here, since nothing downstream can succeed without a valid response).
'''  - REAL BUG FIXED: VBA's "And" is not short-circuiting (unlike AndAlso), so the
'''    original's "If ThirdPartyAccountId = "" And CInt(IOCustNum) < 9000" always
'''    evaluates CInt(IOCustNum) even when IOCustNum is blank - which crashes (CInt("")
'''    is invalid). That means every time IOCustNum comes back empty, the original logs
'''    the correct "-3 missing" error and then immediately crashes again on the very next
'''    line, generating a second spurious log entry. Not reproduced: the second check is
'''    skipped entirely when IOCustNum is blank.
'''  - NOT A BUG, preserved exactly: the INSERT stores IOCustNum (the custom-field value)
'''    into the ThirdPartyAccountId column, and Evo's own native ThirdPartyAccountId field
'''    into the KubeAccountId column. This looks swapped but isn't - confirmed against
'''    Command73_Click's joins (Evo_Customer_XRef.KubeAccountId -> Kube invoices,
'''    Evo_Customer_XRef.ThirdPartyAccountId -> QuickBooks account numbers). Evo's native
'''    ThirdPartyAccountId field means "Kube's account ID, from Evo's perspective" (Kube
'''    being the third party to Evo); IOCustNum is QuickBooks' account number (QuickBooks
'''    being the third party to this application). Two valid "third party" perspectives
'''    colliding in one confusingly-named pair of columns - not touched.
'''  - Per-item failures (missing IOCustNum, missing ThirdPartyAccountId for a likely-real
'''    customer, upsert failure) are logged and counted, but processing still continues to
'''    attempt the upsert with whatever data is available - matching the original's actual
'''    Resume Next behavior (it doesn't skip the upsert after logging these), and matching
'''    your stated general preference: log for review, don't stop the run over one item.
'''  - "Take the LAST CustomFields[].Value" (not matched by field name/key) is preserved
'''    exactly from the original, since I don't have visibility into whether this API ever
'''    returns more than one custom field in practice. Worth validating against real data -
'''    if CustomFields ever contains more than one entry, this could silently pick up the
'''    wrong value.
'''  - Table name verified: Evo_Customer_XRef_SQL -> real name Evo_Customer_XRef (confirmed
'''    via its tbldefs .json).
'''  - Returns an error count instead of MsgBox/Batch-mode messaging - same pattern as
'''    every other job; the caller decides how to surface it.
''' </summary>
Public Module CustomerXrefJob

    Private Const ApiUrl As String = "https://io.hostedsuite.com/api/json/reply/ListClientNamesRequest"

    Public Async Function RunAsync() As Task(Of Integer)
        Dim errorCount = 0
        Dim evoPassword = ConfigHelper.GetConfigValue("Evo Pass")

        Dim payload = New With {.CustomerName = "IO", .UserName = "sanfran", .Password = evoPassword}

        Dim response As ApiClient.ApiResponse
        Try
            response = Await ApiClient.PostAsync(ApiUrl, payload, timeoutSeconds:=60)
            response.EnsureSuccess() ' original: Err.Raise -1 on non-OK
        Catch ex As Exception
            ErrorLogHelper.LogError("Update Customer XRef", "Error retrieving customer from Evo")
            Return 1
        End Try

        Dim items = response.DataAs(Of List(Of EvoClientItem))()

        For Each item In items
            ' original: takes the LAST CustomFields[].Value, not matched by field name - see remarks.
            Dim ioCustNum = item.CustomFields?.LastOrDefault()?.Value

            If String.IsNullOrEmpty(ioCustNum) Then
                ErrorLogHelper.LogError("Update Customer XRef", $"IOCustNum missing in {item.Id}")
                errorCount += 1
                ' original crashes here via CInt("") on the next check - not reproduced, see remarks.
            ElseIf String.IsNullOrEmpty(item.ThirdPartyAccountId) AndAlso Integer.Parse(ioCustNum) < 9000 Then
                ErrorLogHelper.LogError("Update Customer XRef", $"ThirdPartyReference ID missing in {item.Id}")
                errorCount += 1
                ' original still attempts the upsert after logging this - matching that.
            End If

            Try
                UpsertXref(item.Id, item.Name, ioCustNum, item.ThirdPartyAccountId)
            Catch ex As Exception
                ErrorLogHelper.LogError("Update Customer XRef", $"SQL error upserting {item.Id}: {ex.Message}")
                errorCount += 1
            End Try
        Next

        Return errorCount
    End Function

    ''' <summary>
    ''' See class remarks re: the ThirdPartyAccountId/KubeAccountId column mapping - this
    ''' is intentional, not a bug, despite looking swapped at a glance.
    ''' </summary>
    Private Sub UpsertXref(id As String, name As String, ioCustNum As String, evoThirdPartyAccountId As String)
        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            conn.Open()

            Dim existingName As String = Nothing
            Using checkCmd As New SqlCommand("SELECT Name FROM Evo_Customer_XRef WHERE Id = @Id", conn)
                checkCmd.Parameters.AddWithValue("@Id", id)
                Dim result = checkCmd.ExecuteScalar()
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then existingName = result.ToString()
            End Using

            If existingName Is Nothing Then
                Const insertSql As String =
                    "INSERT INTO Evo_Customer_XRef (Id, Name, ThirdPartyAccountId, KubeAccountId) " &
                    "VALUES (@Id, @Name, @ThirdPartyAccountId, @KubeAccountId)"

                Using cmd As New SqlCommand(insertSql, conn)
                    cmd.Parameters.AddWithValue("@Id", id)
                    cmd.Parameters.AddWithValue("@Name", name)
                    cmd.Parameters.AddWithValue("@ThirdPartyAccountId", CType(If(ioCustNum, DBNull.Value), Object))
                    cmd.Parameters.AddWithValue("@KubeAccountId", CType(If(evoThirdPartyAccountId, DBNull.Value), Object))
                    cmd.ExecuteNonQuery()
                End Using
            ElseIf existingName <> name Then
                ' original only ever updates Name on existing rows - ThirdPartyAccountId/
                ' KubeAccountId are set once at insert and never revisited. Preserved as-is;
                ' this looks like a deliberate "these mappings are stable once created"
                ' choice rather than an oversight, but flagging in case it isn't.
                Const updateSql As String = "UPDATE Evo_Customer_XRef SET Name = @Name WHERE Id = @Id"
                Using cmd As New SqlCommand(updateSql, conn)
                    cmd.Parameters.AddWithValue("@Name", name)
                    cmd.Parameters.AddWithValue("@Id", id)
                    cmd.ExecuteNonQuery()
                End Using
            End If
        End Using
    End Sub

End Module
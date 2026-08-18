Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command53_Click() ("RemoteLock Users" button).
''' Delta-syncs Customer_Ops_All against RemoteLock: deletes terminated customers' access,
''' adds new customers, updates changed name/pin/department, and re-provisions door access
''' when membership status changes. Finishes by removing any RemoteLock user with no
''' matching active customer record.
'''
''' Changes from the VBA original - see also remarks on RemoteLockAuth, ApiThrottle,
''' RemoteLockPerson for related notes:
'''  - RemoteLock_Temp was a local Access table (confirmed - no linked-table descriptor,
'''    same as Spheremail_Storage_Temp) used only as an in-run PIN->ID cache. Replaced by
'''    an in-memory Dictionary here - no DB round-trip.
'''  - apicounter is tracked per-phase locally rather than as one value threaded through
'''    the whole routine - Async functions can't take ByRef parameters in VB.NET, so a
'''    single mutable counter can't be passed through the call chain the way VBA's could.
'''    Rate-limiting still triggers roughly every 100 calls, just per-phase instead of
'''    globally across the whole run.
'''  - Each customer's processing logs and counts every expected failure (a non-2xx API
'''    response) at the exact point the original would have hit it, matching the
'''    original's actual Resume Next behavior at each of its 8 relevant error points -
'''    some fall through and keep going (delete failure, PUT failure, per-access delete
'''    failures during re-provisioning), others explicitly GoTo continue and end that
'''    row's processing (add failure, change-detection GET failure, access-list GET
'''    failure, both access-grant POST failures). See remarks on ProcessCustomerRowAsync
'''    for the full mapping. A Try/Catch around each row in RunAsync is kept only for
'''    genuinely unexpected failures (network errors, timeouts, malformed JSON) - auth
'''    failure and the initial full RemoteLock user fetch failing still abort the whole
'''    run rather than continuing with an empty/garbage token, since every subsequent
'''    call would fail identically and just flood the log.
'''  - Mode/completion messaging (Batch <> "X" MsgBox variants) is NOT handled here -
'''    this returns an error count; the caller decides how to surface it, same pattern
'''    as SphereMailStorageEmailJob.
''' </summary>
Public Module RemoteLockUsersJob

    Private Const ApiBaseUrl As String = "https://api.remotelock.com"

    ' Hardcoded RemoteLock accessible_id values from the original - these identify
    ' specific locations/locks/door groups/schedules in RemoteLock's system and can't be
    ' derived from anything else in this codebase.
    Private Const SfLocationId As String = "a0aaaa97-2b30-4515-88a9-4263ccf9bb02"
    Private Const BgLocationId As String = "aab6eb62-c55f-4cd6-bd36-37c31869d012"
    Private Const SfFrontDoorId As String = "cf69b02f-c854-4358-b4c8-eb81a1b492c4"
    Private Const BgFrontDoorId As String = "62392638-862e-4d7c-8add-7234357f9b6c"
    Private Const SfHourlyDoorGroupId As String = "70230b83-b382-43bc-8c06-fddd42262b76"
    Private Const BgHourlyDoorGroupId As String = "9d4c3b15-f6fb-470e-b677-f6463878d7e8"
    Private Const HourlyAccessScheduleId As String = "8df77920-0466-4114-9192-2bb577fbd7e9"

    Private Class CustomerDeltaRow
        Public Property IsMember As Boolean
        Public Property ContactName As String
        Public Property VersionHeader As String
        Public Property VersionCont As String
        Public Property FacilitiesCode As String
        Public Property IsTerminated As Boolean
        Public Property RemoteLockManual As Boolean
        Public Property PrimaryOffice As String
    End Class

    Public Async Function RunAsync() As Task(Of Integer)
        Dim errorCount = 0
        Dim accessToken As String

        ' --- Auth ---
        Try
            Dim tokenResult = Await RemoteLockAuth.RefreshTokenAsync()
            accessToken = tokenResult.AccessToken
        Catch ex As Exception
            ErrorLogHelper.LogError("RemoteLock Users", $"Error retrieving Refresh Token in RemoteLock users process: {ex.Message}")
            Return 1
        End Try

        ' --- Fetch all current RemoteLock users (paginated) ---
        Dim allPersons As List(Of RemoteLockPerson)
        Try
            allPersons = Await FetchAllPersonsAsync(accessToken)
        Catch ex As Exception
            ErrorLogHelper.LogError("RemoteLock Users", $"Error retrieving RemoteLock users in RemoteLock users process: {ex.Message}")
            Return 1
        End Try

        Dim personsByPin As New Dictionary(Of String, RemoteLockPerson)
        For Each p In allPersons
            personsByPin(p.Pin) = p ' last-write-wins on duplicate PINs, same effective behavior as DLookup's non-deterministic first match
        Next

        ' --- Process delta customer rows ---
        Dim lastDelta = ConfigHelper.GetConfigValue("RemoteLock Last Delta")
        Dim deltaRows = GetDeltaCustomerRows(lastDelta)

        For Each row In deltaRows
            If row.VersionHeader > lastDelta Then lastDelta = row.VersionHeader
            If row.VersionCont > lastDelta Then lastDelta = row.VersionCont

            Try
                errorCount += Await ProcessCustomerRowAsync(row, accessToken, personsByPin, allPersons)
            Catch ex As Exception
                ' Genuinely unexpected failures only (network errors, timeouts, malformed
                ' JSON) - expected API failures (non-2xx responses) are logged and counted
                ' inside ProcessCustomerRowAsync itself without throwing. See remarks there.
                ErrorLogHelper.LogError("RemoteLock Users", $"Unexpected error processing {row.ContactName} in RemoteLock users process: {ex.Message}")
                errorCount += 1
            End Try
        Next

        ConfigHelper.SetConfigValue("RemoteLock Last Delta", lastDelta)

        ' --- Orphan cleanup: RemoteLock users with no matching active customer record ---
        errorCount += Await CleanupOrphansAsync(accessToken, allPersons)

        Return errorCount
    End Function

    ''' <summary>
    ''' Processes one customer's delta: delete (if terminated), add (if new), or
    ''' modify name/pin/department and/or door access (if changed). Returns the number of
    ''' errors encountered - every expected failure (a non-2xx API response) is logged and
    ''' counted here directly rather than thrown, matching the original's actual Resume
    ''' Next behavior at each of its 8 relevant error points:
    '''   - Delete failure, PUT (name/pin/dept) failure, and per-access DELETE failures
    '''     during re-provisioning: original falls through and keeps going (no GoTo
    '''     continue) - so those log and continue here too, in the same place.
    '''   - Add failure, change-detection GET failure, access-list GET failure, and both
    '''     access-grant POST failures: original explicitly GoTo continue - so those log
    '''     and Return early here, ending this row's processing (but NOT the whole run).
    ''' The only Try/Catch left is in the caller (RunAsync), reserved for genuinely
    ''' unexpected failures (network errors, timeouts, malformed JSON) rather than
    ''' expected non-2xx responses.
    ''' </summary>
    Private Async Function ProcessCustomerRowAsync(row As CustomerDeltaRow, accessToken As String,
                                                     personsByPin As Dictionary(Of String, RemoteLockPerson),
                                                     allPersons As List(Of RemoteLockPerson)) As Task(Of Integer)
        Dim errorCount = 0
        Dim member = If(row.IsMember, "Member", "")
        Dim existingPerson As RemoteLockPerson = Nothing
        personsByPin.TryGetValue(row.FacilitiesCode, existingPerson)
        Dim id = existingPerson?.Id

        ' --- Delete ---
        If row.IsTerminated AndAlso id IsNot Nothing Then
            Dim headers = AuthHeader(accessToken)
            Dim response = Await ApiClient.DeleteAsync($"{ApiBaseUrl}/access_persons/{id}", headers, timeoutSeconds:=15)
            If response.StatusCode <> Net.HttpStatusCode.NoContent Then
                ErrorLogHelper.LogError("RemoteLock Users", $"Error deleting {row.ContactName} in RemoteLock users process 1st routine")
                errorCount += 1
            End If

            personsByPin.Remove(row.FacilitiesCode)
            allPersons.Remove(existingPerson)
            Return errorCount ' this branch always ends the row's processing either way, same as original
        End If

        ' --- Manual override: skip entirely ---
        If row.RemoteLockManual Then Return 0

        Dim modify As String = "" ' "" none, "X" name/pin/dept, "Y" access only, "Z" both

        ' --- Add ---
        If id Is Nothing AndAlso Not row.IsTerminated AndAlso Not String.IsNullOrEmpty(row.FacilitiesCode) Then
            Dim payload = New With {
                .type = "access_user",
                .attributes = New With {.name = row.ContactName, .pin = row.FacilitiesCode, .department = member}
            }
            Dim headers = AuthHeader(accessToken)
            Dim response = Await ApiClient.PostAsync($"{ApiBaseUrl}/access_persons", payload, headers, timeoutSeconds:=15)

            If response.StatusCode <> Net.HttpStatusCode.Created Then
                ErrorLogHelper.LogError("RemoteLock Users", $"Error adding {row.ContactName} in RemoteLock users process")
                Return errorCount + 1 ' original: GoTo continue
            End If

            id = response.DataAs(Of RemoteLockPersonResponse)().Data.Id
            modify = "Y"
        End If

        ' --- Change detection ---
        If id IsNot Nothing AndAlso Not String.IsNullOrEmpty(row.FacilitiesCode) AndAlso modify = "" Then
            Dim headers = AuthHeader(accessToken)
            Dim response = Await ApiClient.GetAsync($"{ApiBaseUrl}/access_persons/{id}", Nothing, headers, timeoutSeconds:=15)

            If Not response.IsSuccess Then
                ErrorLogHelper.LogError("RemoteLock Users", $"Error retrieving details for {row.ContactName} in RemoteLock users process")
                Return errorCount + 1 ' original: GoTo continue
            End If

            Dim current = response.DataAs(Of RemoteLockPersonResponse)().Data.Attributes

            If current.Name <> row.ContactName Then modify = "X"
            If current.Pin <> row.FacilitiesCode Then modify = "X"
            If member <> current.Department Then modify = "Z"
        End If

        ' --- Name/PIN/Department update ---
        If modify = "X" Or modify = "Z" Then
            Dim payload = New With {
                .attributes = New With {.name = row.ContactName, .pin = row.FacilitiesCode, .department = member}
            }
            Dim headers = AuthHeader(accessToken)
            Dim response = Await ApiClient.PutAsync($"{ApiBaseUrl}/access_persons/{id}", payload, headers, timeoutSeconds:=15)

            If Not response.IsSuccess Then
                ' original does NOT skip the access-grant step below on this failure - falls through, so we do too.
                ErrorLogHelper.LogError("RemoteLock Users", $"Error modifying details for {row.ContactName} in RemoteLock users process")
                errorCount += 1
            End If
        End If

        ' --- Access re-provisioning ---
        If modify = "Y" Or modify = "Z" Then
            Dim headers = AuthHeader(accessToken)

            Dim existingAccessResponse = Await ApiClient.GetAsync($"{ApiBaseUrl}/access_persons/{id}/accesses", Nothing, headers, timeoutSeconds:=15)
            If Not existingAccessResponse.IsSuccess Then
                ErrorLogHelper.LogError("RemoteLock Users", $"Error retrieving current access for {row.ContactName} in RemoteLock users process")
                Return errorCount + 1 ' original: GoTo continue
            End If

            Dim existingAccesses = existingAccessResponse.DataAs(Of RemoteLockAccessListResponse)().Data
            For Each access In existingAccesses
                Dim deleteResponse = Await ApiClient.DeleteAsync($"{ApiBaseUrl}/access_persons/{id}/accesses/{access.Id}", headers, timeoutSeconds:=15)
                If deleteResponse.StatusCode <> Net.HttpStatusCode.NoContent Then
                    ErrorLogHelper.LogError("RemoteLock Users", $"Error removing current access for {row.ContactName} in RemoteLock users process")
                    errorCount += 1
                End If
            Next

            Dim isSf = row.PrimaryOffice = "San Francisco"
            Dim primaryPayload As Object

            If member = "Member" Then
                primaryPayload = New With {.attributes = New With {.accessible_id = If(isSf, SfLocationId, BgLocationId), .accessible_type = "location"}}
            Else
                primaryPayload = New With {.attributes = New With {.accessible_id = If(isSf, SfFrontDoorId, BgFrontDoorId), .accessible_type = "lock"}}
            End If

            Dim primaryResponse = Await ApiClient.PostAsync($"{ApiBaseUrl}/access_persons/{id}/accesses", primaryPayload, headers, timeoutSeconds:=15)
            If primaryResponse.StatusCode <> Net.HttpStatusCode.Created Then
                ErrorLogHelper.LogError("RemoteLock Users", $"Error adding new access for {row.ContactName} in RemoteLock users process")
                Return errorCount + 1 ' original: GoTo continue
            End If

            If member <> "Member" Then
                Dim doorGroupPayload = New With {
                    .attributes = New With {
                        .accessible_id = If(isSf, SfHourlyDoorGroupId, BgHourlyDoorGroupId),
                        .access_schedule_id = HourlyAccessScheduleId,
                        .accessible_type = "door_group"
                    }
                }
                Dim doorGroupResponse = Await ApiClient.PostAsync($"{ApiBaseUrl}/access_persons/{id}/accesses", doorGroupPayload, headers, timeoutSeconds:=15)
                If doorGroupResponse.StatusCode <> Net.HttpStatusCode.Created Then
                    ErrorLogHelper.LogError("RemoteLock Users", $"Error adding door group access for {row.ContactName} in RemoteLock users process")
                    errorCount += 1 ' last possible step either way, so a plain increment here is equivalent to "GoTo continue"
                End If
            End If
        End If

        Return errorCount
    End Function

    Private Async Function FetchAllPersonsAsync(accessToken As String) As Task(Of List(Of RemoteLockPerson))
        Dim result As New List(Of RemoteLockPerson)
        Dim headers = AuthHeader(accessToken)
        Dim page = 1
        Dim totalPages = 1
        Dim throttleCounter = 0

        While page <= totalPages
            Dim queryParams = New Dictionary(Of String, String) From {{"page", page.ToString()}}
            Dim response = Await ApiClient.GetAsync($"{ApiBaseUrl}/access_persons", queryParams, headers, timeoutSeconds:=15)

            If Not response.IsSuccess Then
                Throw New InvalidOperationException($"RemoteLock access_persons page {page} returned {CInt(response.StatusCode)}")
            End If

            Dim data = response.DataAs(Of RemoteLockPersonListResponse)()
            If page = 1 Then totalPages = data.Meta.TotalPages

            For Each item In data.Data
                If item.Type = "access_user" Then
                    result.Add(New RemoteLockPerson With {
                        .Id = item.Id,
                        .AccessName = item.Attributes.Name,
                        .Pin = item.Attributes.Pin,
                        .Status = item.Attributes.Status,
                        .Department = item.Attributes.Department
                    })
                End If
            Next

            throttleCounter = Await ApiThrottle.ThrottleIfNeededAsync(throttleCounter)
            page += 1
        End While

        Return result
    End Function

    ''' <summary>
    ''' Removes RemoteLock users with no matching active customer record (Facilities Code
    ''' no longer found in Customer_Ops), excluding internal/staff accounts.
    ''' </summary>
    Private Async Function CleanupOrphansAsync(accessToken As String, allPersons As List(Of RemoteLockPerson)) As Task(Of Integer)
        Dim errorCount = 0
        Dim activeCodes = GetActiveFacilitiesCodes()
        Dim headers = AuthHeader(accessToken)
        Dim throttleCounter = 0

        Dim orphans = allPersons.
            Where(Function(p) Not activeCodes.Contains(p.Pin) AndAlso p.Department <> "Internal").
            ToList()

        For Each orphan In orphans
            Try
                Dim response = Await ApiClient.DeleteAsync($"{ApiBaseUrl}/access_persons/{orphan.Id}", headers, timeoutSeconds:=15)
                If response.StatusCode <> Net.HttpStatusCode.NoContent Then
                    ErrorLogHelper.LogError("RemoteLock Users", $"Error deleting {orphan.AccessName} in RemoteLock users process 2nd routine")
                    errorCount += 1
                End If
                throttleCounter = Await ApiThrottle.ThrottleIfNeededAsync(throttleCounter)
            Catch ex As Exception
                ErrorLogHelper.LogError("RemoteLock Users", $"Error deleting {orphan.AccessName} in RemoteLock users process 2nd routine: {ex.Message}")
                errorCount += 1
            End Try
        Next

        Return errorCount
    End Function

    Private Function AuthHeader(accessToken As String) As Dictionary(Of String, String)
        Return New Dictionary(Of String, String) From {{"Authorization", $"Bearer {accessToken}"}}
    End Function

    Private Function GetDeltaCustomerRows(lastDelta As String) As List(Of CustomerDeltaRow)
        Dim result As New List(Of CustomerDeltaRow)
        Const sql As String =
            "SELECT [Is Member], [Contact Name], [Version Header], [Version Cont], [Facilities Code], " &
            "Terminated, Terminated_Cont, RemoteLock_Manual, [Primary Office] " &
            "FROM Customer_Ops_All " &
            "WHERE [Version Header] > @LastDelta OR [Version Cont] > @LastDelta"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@LastDelta", lastDelta)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result.Add(New CustomerDeltaRow With {
                            .IsMember = TryGetString(reader, "Is Member") = "X",
                            .ContactName = TryGetString(reader, "Contact Name"),
                            .VersionHeader = TryGetString(reader, "Version Header"),
                            .VersionCont = TryGetString(reader, "Version Cont"),
                            .FacilitiesCode = TryGetString(reader, "Facilities Code"),
                            .IsTerminated = Not reader.IsDBNull(reader.GetOrdinal("Terminated")) OrElse Not reader.IsDBNull(reader.GetOrdinal("Terminated_Cont")),
                            .RemoteLockManual = TryGetString(reader, "RemoteLock_Manual") = "X",
                            .PrimaryOffice = TryGetString(reader, "Primary Office")
                        })
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Function GetActiveFacilitiesCodes() As HashSet(Of String)
        Dim result As New HashSet(Of String)
        Const sql As String = "SELECT [Facilities Code] FROM Customer_Ops WHERE [Facilities Code] IS NOT NULL AND [Facilities Code] <> ''"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result.Add(reader.GetString(0))
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Function TryGetString(reader As SqlDataReader, columnName As String) As String
        Dim ordinal = reader.GetOrdinal(columnName)
        Return If(reader.IsDBNull(ordinal), String.Empty, reader.GetValue(ordinal).ToString())
    End Function

End Module
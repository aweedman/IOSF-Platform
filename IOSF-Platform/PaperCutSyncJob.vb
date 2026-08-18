Imports Microsoft.Data.SqlClient

''' <summary>
''' Direct port of Landing Page.cls: Command51_Click() ("Papercut Scan Actions and Users").
''' Three jobs in one routine:
'''   1. Sync scan-to-email actions for eligible customers (create/undelete/update).
'''   2. Soft-delete scan actions for customers who no longer qualify.
'''   3. Sync PaperCut login activation state against Customer_Ops_All/IO_Employees.
'''
''' Changes from the VBA original:
'''  - CONFIRMED WITH AL, kept as-is: this function aborts entirely on its first error
'''    (no Resume Next in the original, unlike every other ported function) - a single
'''    Try/Catch around the whole run, matching that.
'''  - Papercut tables live in a different database on the same SQL Server instance as
'''    everything else (confirmed with Al) - see ConfigHelper.GetPapercutConnectionString.
'''  - Table names verified against tbldefs, NOT simple suffix-strips: scan_action_SQL ->
'''    dbo.tbl_scan_action, scan_action_attribute_SQL -> dbo.tbl_scan_action_attribute,
'''    scan_action_group_SQL -> dbo.tbl_scan_action_group (tbl_user_SQL -> tbl_user happens
'''    to look similar but is a coincidence, not a pattern to assume elsewhere).
'''  - The "sf@ioprint.me" From-address override is now read from Config instead of
'''    hardcoded, per the standing no-hardcoded-emails rule.
'''  - Burlingame is no longer an active location (per instruction) - its Case branch was
'''    removed entirely, rather than kept as a second valid office. Any customer whose
'''    Primary Office still shows "Burlingame" (e.g. an undeleted record whose Customer_Ops
'''    entry hasn't caught up yet) now falls into the same "unrecognized office" skip/log
'''    path as any other unexpected value, and no longer sends anything to
'''    burlingame@ioprint.me. If that undeleted customer's office gets corrected to "San
'''    Francisco" in Customer_Ops on a later run, the normal undelete-then-sync flow
'''    (UndeleteScanAction falls through to SyncGroupIfChanged/SyncFromEmailIfChanged)
'''    already handles moving their group_id and from-email over to SF correctly - no
'''    special-casing needed for that transition.
'''  - The account-key encoding (AccountNumber + Cont_Num concatenated, later parsed back
'''    via Left(4)/Mid(5)) assumes Account Number is always exactly 4 characters - this
'''    assumption comes from the original, not introduced here. If that's ever untrue,
'''    the parse in DeactivateStaleScanActions will silently misread the account number.
''' </summary>
Public Module PaperCutSyncJob

    Private Const TemplateScanActionId As Integer = 3003 ' hardcoded template row copied for new scan actions
    Private Const ModifiedByUserId As String = "77005" ' hardcoded "system" user id in the original

    Public Function Run() As Integer
        Try
            Using conn As New SqlConnection(ConfigHelper.GetPapercutConnectionString())
                conn.Open()
                SyncScanActionsForEligibleCustomers(conn)
                DeactivateStaleScanActions(conn)
            End Using

            SyncUserActivation()
            Return 0
        Catch ex As Exception
            ' Matches original: logs exactly one error and stops - no Resume Next here,
            ' confirmed intentional to preserve.
            ErrorLogHelper.LogError("Papercut Scan Actions", $"SQL error: {ex.Message}")
            Return 1
        End Try
    End Function

    Private Sub SyncScanActionsForEligibleCustomers(conn As SqlConnection)
        Dim customers = GetEligibleCustomers()
        Dim nextIdHint = 1001

        For Each customer In customers
            Dim account = $"{customer.AccountNumber}{customer.ContNum}"
            Dim label = BuildLabel(customer.ContactName)
            If customer.FrequentScans Then label = "." & label ' sorts frequent scanners to top of the list

            Dim current = DateTime.Now
            Dim ticks = BuildTicks(current)

            Dim groupId As Integer
            Dim fromEmailConfigKey As String
            Select Case customer.PrimaryOffice
                Case "San Francisco"
                    groupId = 4009
                    fromEmailConfigKey = "Papercut SF From Email"
                Case Else
                    ' Burlingame is no longer an active location (removed per instruction) -
                    ' any customer still showing "Burlingame" here (e.g. an undeleted record
                    ' whose Customer_Ops entry hasn't been updated) is now treated as an
                    ' anomaly rather than a valid office, and skipped/logged like any other
                    ' unrecognized value instead of silently emailing burlingame@ioprint.me.
                    ErrorLogHelper.LogError("Papercut Scan Actions", $"Skipped {customer.ContactName} ({account}) - unrecognized or inactive Primary Office '{customer.PrimaryOffice}'")
                    Continue For
            End Select
            Dim fromEmail = ConfigHelper.GetConfigValue(fromEmailConfigKey)

            nextIdHint += 1

            Dim existingId = GetExistingScanActionId(conn, account)

            If existingId Is Nothing Then
                Dim newId = FindNextAvailableId(conn, nextIdHint)
                nextIdHint = newId
                CreateNewScanAction(conn, newId, label, customer.EmailAddress, ticks, current, account, groupId)
                Continue For ' original: GoTo skip - freshly created, nothing else to sync
            ElseIf IsMarkedDeleted(conn, existingId.Value) Then
                UndeleteScanAction(conn, existingId.Value, label, ticks, current)
                ' original falls through to the sync checks below even after undeleting - matching that.
            End If

            SyncLabelIfChanged(conn, existingId.Value, label, ticks, current)
            SyncGroupIfChanged(conn, existingId.Value, groupId, ticks, current)
            SyncFromEmailIfChanged(conn, existingId.Value, fromEmail, ticks)
            SyncTargetEmailIfChanged(conn, existingId.Value, customer.EmailAddress, ticks, current)
        Next
    End Sub

    ''' <summary>
    ''' Reformats "First Last" (or "First Middle Last") to "Last, First..." by splitting on
    ''' the FIRST space only - matches the original's Left/Right-based split exactly,
    ''' including its limitation with multi-word names.
    '''
    ''' REAL BUG FIXED: this used to also escape single quotes here (.Replace("'", "''")),
    ''' but the result is used as a parameterized @Label value in most call sites -
    ''' parameterized values must NEVER be pre-escaped, since ADO.NET sends them to SQL
    ''' Server literally with no text interpretation. Escaping here meant the doubled
    ''' apostrophe got stored as literal data - confirmed via PaperCut's own UI showing
    ''' "O''Neil" instead of "O'Neil" for a real customer. The one place label is compared
    ''' via raw concatenated SQL text (SyncLabelIfChanged's WHERE clause) already applies
    ''' its own escaping via ToSqlLiteral() - that was being doubled by this redundant one.
    ''' </summary>
    Private Function BuildLabel(contactName As String) As String
        Dim spacePos = contactName.IndexOf(" "c)
        Return If(spacePos >= 0,
            contactName.Substring(spacePos + 1) & ", " & contactName.Substring(0, spacePos),
            contactName)
    End Function

    Private Function BuildTicks(current As DateTime) As String
        Dim secondsSinceEpoch = CLng((current - New DateTime(1970, 1, 1)).TotalSeconds)
        Return $"{secondsSinceEpoch}000"
    End Function

    Private Function GetEligibleCustomers() As List(Of ScanActionCustomer)
        Dim result As New List(Of ScanActionCustomer)
        Const sql As String =
            "SELECT [Contact Name], [Primary Office], [Email Address], Frequent_Scans, [Account Number], Cont_Num " &
            "FROM Customer_Ops " &
            "WHERE ([Service Level] LIKE '%MAIL%' OR [Service Level] LIKE '%EXEC%' OR [Service Level] LIKE '%BIS%' OR " &
            "[Service Level] LIKE '%BIF%' OR [Service Level] LIKE '%BIL%') AND " &
            "[Service Level] NOT LIKE '%VBIS%' AND [Email Address] IS NOT NULL AND [Email Address] <> ''"

        Using conn As New SqlConnection(ConfigHelper.ConnectionString)
            Using cmd As New SqlCommand(sql, conn)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        result.Add(New ScanActionCustomer With {
                            .ContactName = reader.GetString(0),
                            .PrimaryOffice = If(reader.IsDBNull(1), String.Empty, reader.GetString(1)),
                            .EmailAddress = reader.GetString(2),
                            .FrequentScans = Not reader.IsDBNull(3) AndAlso reader.GetString(3) = "X",
                            .AccountNumber = reader.GetValue(4).ToString(),
                            .ContNum = reader.GetValue(5).ToString()
                        })
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Function GetExistingScanActionId(conn As SqlConnection, account As String) As Integer?
        Const sql As String =
            "SELECT scan_action_id FROM tbl_scan_action_attribute WHERE attrib_value = @Account AND attrib_name = 'notes'"
        Using cmd As New SqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@Account", account)
            Dim result = cmd.ExecuteScalar()
            Return If(result Is Nothing OrElse result Is DBNull.Value, CType(Nothing, Integer?), Convert.ToInt32(result))
        End Using
    End Function

    Private Function FindNextAvailableId(conn As SqlConnection, startId As Integer) As Integer
        Dim candidateId = startId
        While CountRows(conn, "tbl_scan_action", "scan_action_id = @Id", candidateId) > 0 OrElse
              CountRows(conn, "tbl_scan_action_group", "scan_action_group_id = @Id", candidateId) > 0
            candidateId += 1
        End While
        Return candidateId
    End Function

    Private Function CountRows(conn As SqlConnection, table As String, whereClause As String, id As Integer) As Integer
        Using cmd As New SqlCommand($"SELECT COUNT(*) FROM {table} WHERE {whereClause}", conn)
            cmd.Parameters.AddWithValue("@Id", id)
            Return CInt(cmd.ExecuteScalar())
        End Using
    End Function

    ''' <summary>
    ''' REAL BUG FIXED, not just re-verified - see conversation this was fixed in. The
    ''' original version of this check (in DeactivateStaleScanActions) concatenated
    ''' accountNumber/contNum directly into the WHERE clause with NO escaping and NO
    ''' parameterization at all - unlike every other query in this file, which either uses
    ''' bound parameters or the ToSqlLiteral() escaping helper. accountNumber/contNum come
    ''' from splitting a stored "notes" field by fixed character position (Substring(0,4) /
    ''' Substring(4)) - a position that's only correct if Account Number is always exactly
    ''' 4 characters (see class remarks). If that assumption is ever wrong for some
    ''' customer, whatever falls out of the split gets dropped straight into raw SQL text,
    ''' which can break SQL syntax outright rather than just failing to match a row (this is
    ''' what actually happened - confirmed via a real "Unclosed quotation mark" error).
    ''' Proper parameterization closes this off entirely: the split may still misread the
    ''' account number if the 4-character assumption is wrong, but whatever value results
    ''' can never be interpreted as SQL syntax.
    ''' </summary>
    Private Function IsCustomerStillEligible(conn As SqlConnection, accountNumber As String, contNum As String) As Boolean
        Const sql As String =
            "SELECT COUNT(*) FROM Customer_Ops WHERE [Account Number] = @AccountNumber AND Cont_Num = @ContNum AND " &
            "([Service Level] LIKE '%MAIL%' OR [Service Level] LIKE '%EXEC%' OR [Service Level] LIKE '%BIS%' OR " &
            "[Service Level] LIKE '%BIF%' OR [Service Level] LIKE '%BIL%') AND [Service Level] NOT LIKE '%VBIS%'"
        Using cmd As New SqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@AccountNumber", accountNumber)
            cmd.Parameters.AddWithValue("@ContNum", contNum)
            Return CInt(cmd.ExecuteScalar()) = 1
        End Using
    End Function

    Private Sub CreateNewScanAction(conn As SqlConnection, id As Integer, label As String, email As String,
                                     ticks As String, current As DateTime, account As String, groupId As Integer)
        Using cmd As New SqlCommand(
            "INSERT INTO tbl_scan_action (scan_action_id, action_type, label, target, deleted, modified_ticks, " &
            "modified_date, modified_by, created_date, created_by, deleted_date) " &
            "VALUES (@Id, 'email_to', @Label, @Target, 'N', @Ticks, @Current, @ModifiedBy, @Current, @ModifiedBy, NULL)", conn)
            cmd.Parameters.AddWithValue("@Id", id)
            cmd.Parameters.AddWithValue("@Label", label)
            cmd.Parameters.AddWithValue("@Target", email)
            cmd.Parameters.AddWithValue("@Ticks", ticks)
            cmd.Parameters.AddWithValue("@Current", current)
            cmd.Parameters.AddWithValue("@ModifiedBy", ModifiedByUserId)
            cmd.ExecuteNonQuery()
        End Using

        Using cmd As New SqlCommand(
            "INSERT INTO tbl_scan_action_attribute (scan_action_id, attrib_value, modified_ticks, propagate, attrib_name) " &
            "SELECT @Id, attrib_value, @Ticks, propagate, attrib_name FROM tbl_scan_action_attribute WHERE scan_action_id = @TemplateId", conn)
            cmd.Parameters.AddWithValue("@Id", id)
            cmd.Parameters.AddWithValue("@Ticks", ticks)
            cmd.Parameters.AddWithValue("@TemplateId", TemplateScanActionId)
            cmd.ExecuteNonQuery()
        End Using

        Using cmd As New SqlCommand(
            "UPDATE tbl_scan_action_attribute SET attrib_value = @Account WHERE attrib_name = 'notes' AND scan_action_id = @Id", conn)
            cmd.Parameters.AddWithValue("@Account", account)
            cmd.Parameters.AddWithValue("@Id", id)
            cmd.ExecuteNonQuery()
        End Using

        Using cmd As New SqlCommand(
            "INSERT INTO tbl_scan_action_group (scan_action_group_id, group_id, scan_action_id, modified_ticks, " &
            "modified_date, modified_by, created_date, created_by) " &
            "VALUES (@Id, @GroupId, @Id, @Ticks, @Current, @ModifiedBy, @Current, @ModifiedBy)", conn)
            cmd.Parameters.AddWithValue("@Id", id)
            cmd.Parameters.AddWithValue("@GroupId", groupId)
            cmd.Parameters.AddWithValue("@Ticks", ticks)
            cmd.Parameters.AddWithValue("@Current", current)
            cmd.Parameters.AddWithValue("@ModifiedBy", ModifiedByUserId)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Function IsMarkedDeleted(conn As SqlConnection, scanActionId As Integer) As Boolean
        Return CountRows(conn, "tbl_scan_action", "deleted = 'Y' AND scan_action_id = @Id", scanActionId) > 0
    End Function

    Private Sub UndeleteScanAction(conn As SqlConnection, scanActionId As Integer, label As String, ticks As String, current As DateTime)
        Using cmd As New SqlCommand(
            "UPDATE tbl_scan_action SET deleted = 'N', modified_ticks = @Ticks, modified_date = @Current, " &
            "modified_by = @ModifiedBy, deleted_date = NULL, label = @Label WHERE scan_action_id = @Id", conn)
            cmd.Parameters.AddWithValue("@Ticks", ticks)
            cmd.Parameters.AddWithValue("@Current", current)
            cmd.Parameters.AddWithValue("@ModifiedBy", ModifiedByUserId)
            cmd.Parameters.AddWithValue("@Label", label)
            cmd.Parameters.AddWithValue("@Id", scanActionId)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub SyncLabelIfChanged(conn As SqlConnection, scanActionId As Integer, label As String, ticks As String, current As DateTime)
        If CountRows(conn, "tbl_scan_action", "scan_action_id = @Id AND label = " & ToSqlLiteral(label), scanActionId) = 0 Then
            Using cmd As New SqlCommand(
                "UPDATE tbl_scan_action SET label = @Label, modified_ticks = @Ticks, modified_date = @Current, modified_by = @ModifiedBy WHERE scan_action_id = @Id", conn)
                cmd.Parameters.AddWithValue("@Label", label)
                cmd.Parameters.AddWithValue("@Ticks", ticks)
                cmd.Parameters.AddWithValue("@Current", current)
                cmd.Parameters.AddWithValue("@ModifiedBy", ModifiedByUserId)
                cmd.Parameters.AddWithValue("@Id", scanActionId)
                cmd.ExecuteNonQuery()
            End Using
        End If
    End Sub

    Private Sub SyncGroupIfChanged(conn As SqlConnection, scanActionId As Integer, groupId As Integer, ticks As String, current As DateTime)
        If CountRows(conn, "tbl_scan_action_group", "scan_action_id = @Id AND group_id = " & groupId, scanActionId) = 0 Then
            Using cmd As New SqlCommand(
                "UPDATE tbl_scan_action_group SET group_id = @GroupId, modified_ticks = @Ticks, modified_date = @Current, modified_by = @ModifiedBy WHERE scan_action_id = @Id", conn)
                cmd.Parameters.AddWithValue("@GroupId", groupId)
                cmd.Parameters.AddWithValue("@Ticks", ticks)
                cmd.Parameters.AddWithValue("@Current", current)
                cmd.Parameters.AddWithValue("@ModifiedBy", ModifiedByUserId)
                cmd.Parameters.AddWithValue("@Id", scanActionId)
                cmd.ExecuteNonQuery()
            End Using
        End If
    End Sub

    Private Sub SyncFromEmailIfChanged(conn As SqlConnection, scanActionId As Integer, fromEmail As String, ticks As String)
        If CountRows(conn, "tbl_scan_action_attribute", "scan_action_id = @Id AND attrib_value = " & ToSqlLiteral(fromEmail) & " AND attrib_name = 'mail.from.override'", scanActionId) = 0 Then
            Using cmd As New SqlCommand(
                "UPDATE tbl_scan_action_attribute SET attrib_value = @FromEmail, modified_ticks = @Ticks WHERE scan_action_id = @Id AND attrib_name = 'mail.from.override'", conn)
                cmd.Parameters.AddWithValue("@FromEmail", fromEmail)
                cmd.Parameters.AddWithValue("@Ticks", ticks)
                cmd.Parameters.AddWithValue("@Id", scanActionId)
                cmd.ExecuteNonQuery()
            End Using
        End If
    End Sub

    Private Sub SyncTargetEmailIfChanged(conn As SqlConnection, scanActionId As Integer, email As String, ticks As String, current As DateTime)
        If CountRows(conn, "tbl_scan_action", "scan_action_id = @Id AND target = " & ToSqlLiteral(email), scanActionId) = 0 Then
            Using cmd As New SqlCommand(
                "UPDATE tbl_scan_action SET target = @Target, modified_ticks = @Ticks, modified_date = @Current, modified_by = @ModifiedBy WHERE scan_action_id = @Id", conn)
                cmd.Parameters.AddWithValue("@Target", email)
                cmd.Parameters.AddWithValue("@Ticks", ticks)
                cmd.Parameters.AddWithValue("@Current", current)
                cmd.Parameters.AddWithValue("@ModifiedBy", ModifiedByUserId)
                cmd.Parameters.AddWithValue("@Id", scanActionId)
                cmd.ExecuteNonQuery()
            End Using
        End If
    End Sub

    ''' <summary>
    ''' CountRows takes a raw WHERE fragment for the "has this value already" checks above
    ''' since the comparison value needs to sit inside the SAME query as a parameter would,
    ''' but building a dynamic parameter list per call site added more complexity than it's
    ''' worth here - this escapes single quotes properly, matching the original's Replace
    ''' pattern, so it's not a SQL injection risk despite not being a bound parameter.
    ''' </summary>
    Private Function ToSqlLiteral(value As String) As String
        Return "'" & value.Replace("'", "''") & "'"
    End Function

    ''' <summary>
    ''' Soft-deletes scan actions for customers who no longer qualify. Parses the account
    ''' key back out of the stored "notes" attribute (Left(4)/Mid(5)) - see class remarks
    ''' re: the 4-character Account Number assumption.
    ''' </summary>
    Private Sub DeactivateStaleScanActions(conn As SqlConnection)
        Const selectSql As String =
            "SELECT t1.scan_action_id, t2.attrib_value " &
            "FROM tbl_scan_action AS t1 " &
            "INNER JOIN tbl_scan_action_attribute AS t2 ON t1.scan_action_id = t2.scan_action_id " &
            "WHERE t1.deleted = 'N' AND LEFT(t1.label, 1) <> '""' AND LEFT(t1.label, 1) <> '''' AND t2.attrib_name = 'notes'"

        Dim rows As New List(Of (ScanActionId As Integer, AttribValue As String))
        Using cmd As New SqlCommand(selectSql, conn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    ' REAL BUG FIXED: reader.GetInt32() requires the underlying SQL column
                    ' to be exactly INT and throws InvalidCastException on DECIMAL/NUMERIC -
                    ' confirmed this is scan_action_id's actual type in PaperCut's own
                    ' third-party schema (not something under our control). Convert.ToInt32
                    ' handles whatever numeric type the column actually is.
                    rows.Add((Convert.ToInt32(reader.GetValue(0)), reader.GetString(1)))
                End While
            End Using
        End Using

        For Each row In rows
            If row.AttribValue.Length < 4 Then Continue For ' can't parse a 4-char account number out of this
            Dim accountNumber = row.AttribValue.Substring(0, 4)
            Dim contNum = row.AttribValue.Substring(4)

            Dim stillEligible As Boolean
            Using opsConn = GetOpsConnection()
                stillEligible = IsCustomerStillEligible(opsConn, accountNumber, contNum)
            End Using

            If Not stillEligible Then
                Dim current = DateTime.Now
                Dim ticks = BuildTicks(current)
                Using cmd As New SqlCommand(
                    "UPDATE tbl_scan_action SET deleted = 'Y', modified_ticks = @Ticks, modified_date = @Current, " &
                    "modified_by = @ModifiedBy, deleted_date = @Current WHERE scan_action_id = @Id", conn)
                    cmd.Parameters.AddWithValue("@Ticks", ticks)
                    cmd.Parameters.AddWithValue("@Current", current)
                    cmd.Parameters.AddWithValue("@ModifiedBy", ModifiedByUserId)
                    cmd.Parameters.AddWithValue("@Id", row.ScanActionId)
                    cmd.ExecuteNonQuery()
                End Using
            End If
        Next
    End Sub

    ''' <summary>
    ''' Short-lived connection to the Staging database, used only for the eligibility
    ''' re-check inside DeactivateStaleScanActions (that query needs Customer_Ops, which
    ''' lives in Staging, not Papercut).
    ''' </summary>
    Private Function GetOpsConnection() As SqlConnection
        Dim conn As New SqlConnection(ConfigHelper.ConnectionString)
        conn.Open()
        Return conn
    End Function

    Private Sub SyncUserActivation()
        Using conn As New SqlConnection(ConfigHelper.GetPapercutConnectionString())
            conn.Open()

            ' Deactivate: terminated customers whose PaperCut login is still active.
            ' NOTE: this UPDATE joins across the Papercut and Staging databases in the
            ' original (tbl_user_SQL / Customer_Ops_All_SQL are two separately-linked
            ' tables, but Access's query engine could join them transparently). SQL Server
            ' can't join across databases this way without either a linked server or
            ' three-part naming (Staging.dbo.Customer_Ops_All) - using three-part naming
            ' here, which requires both databases to be on the SAME server (confirmed).
            Using cmd As New SqlCommand(
                "UPDATE t2 SET t2.disabled_printing = 'Y' " &
                "FROM tbl_user AS t2 " &
                "INNER JOIN Staging.dbo.Customer_Ops_All AS t1 ON t1.[Facilities Code] = t2.user_name " &
                "WHERE (t1.Terminated IS NOT NULL OR t1.Terminated_Cont IS NOT NULL) AND t2.disabled_printing = 'N'", conn)
                cmd.ExecuteNonQuery()
            End Using

            ' Reactivate: non-terminated customers whose PaperCut login is currently disabled.
            Using cmd As New SqlCommand(
                "UPDATE t2 SET t2.disabled_printing = 'N' " &
                "FROM tbl_user AS t2 " &
                "INNER JOIN Staging.dbo.Customer_Ops_All AS t1 ON t1.[Facilities Code] = t2.user_name " &
                "WHERE t1.Terminated IS NULL AND t1.Terminated_Cont IS NULL AND t2.disabled_printing = 'Y'", conn)
                cmd.ExecuteNonQuery()
            End Using

            ' Deactivate: active logins with no matching customer record AND no matching employee record.
            Using cmd As New SqlCommand(
                "UPDATE t1 SET t1.disabled_printing = 'Y' " &
                "FROM tbl_user AS t1 " &
                "LEFT JOIN Staging.dbo.Customer_Ops_All AS t2 ON t1.user_name = t2.[Facilities Code] " &
                "LEFT JOIN Staging.dbo.IO_Employees AS t3 ON t1.user_name = t3.user_name " &
                "WHERE t1.disabled_printing = 'N' AND t2.[Facilities Code] IS NULL AND t3.user_name IS NULL", conn)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

End Module
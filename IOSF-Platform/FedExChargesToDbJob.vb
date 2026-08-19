Imports Microsoft.Data.SqlClient
Imports CsvHelper
Imports System.Globalization
Imports System.IO

''' <summary>
''' Direct port of Landing Page.cls: Command50_Click() ("160.2 - FedEx Charges to DB").
'''
''' Account filter (column 2 = "324282815") preserved exactly as a hardcoded literal,
''' matching the original - the raw CSV export apparently combines charges across
''' multiple FedEx accounts, and only this specific one is relevant here. Not otherwise
''' explained in the source, so not renamed/reinterpreted.
'''
''' Date parsing (column 3): the original reassembles an 8-character YYYYMMDD string into
''' MM/DD/YYYY via Mid/Right/Left string slicing - replicated the same way (Year =
''' chars[0:4], Month = chars[4:6], Day = chars[6:8]), rather than assuming a different
''' format.
'''
''' AGGREGATION LOGIC preserved exactly, including a real edge case in the original's own
''' design: it checks whether the LOOKED-UP Total_Cost equals 0 to decide INSERT vs
''' UPDATE, not whether a matching row actually exists. If an existing row's Total_Cost
''' happens to be exactly 0 already, the original would attempt a duplicate INSERT rather
''' than an UPDATE - and the original's own error handler specifically suppresses Access
''' error 3022 (duplicate key violation) without logging it, silently skipping to the
''' next row. This is replicated as-is: SQL Server's equivalent duplicate-key error
''' numbers (2627, 2601) are caught and silently skipped, matching the original's
''' deliberate suppression, rather than "fixed" to check row existence more robustly -
''' this is the original's own intentional design, not an accident.
'''
''' File picker: the original's AllowMultiSelect=True is unused (only .SelectedItems(1)
''' is ever read), same dead-multi-select pattern already seen in SendPro Forwards - this
''' uses a plain single-file picker matching actual behavior.
'''
''' DELETE-BEFORE-INSERT ADDED per Al (not in the original): before processing any rows,
''' deletes existing FedEx rows for the selected Billing_Start_Date - makes re-running/
''' reloading the same file safe without depending solely on the insert/update
''' aggregation logic below. This does NOT wrap the whole run in one atomic transaction -
''' the delete is its own quick upfront step, and the per-row processing loop keeps its
''' existing per-row error handling (a failure on one row doesn't roll back or stop the
''' rest), matching the original's own per-row resilience design.
'''
''' Table name NOT independently verified: FedEx_SQL -> assumed real name FedEx (the
''' simple-strip convention used elsewhere in this port, but not confirmed against a
''' tbldefs descriptor).
''' </summary>
Public Module FedExChargesToDbJob

    Private Const TargetAccountNumber As String = "324282815"

    Public Function RunAsync(csvFilePath As String, billStartDate As Date) As Task(Of Integer)
        Return Task.Run(Function()
                            Dim errorCount = 0

                            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                                conn.Open()

                                ' DELETE-BEFORE-INSERT ADDED per Al (not in the original), same pattern
                                ' as SendProForwardsToDbJob/KubeMeetingsToDbJob: clears existing rows for
                                ' this Billing_Start_Date before processing, so the file can be safely
                                ' re-run/reloaded without relying on the insert/update aggregation logic
                                ' alone. As a side effect, this also makes the Total_Cost=0 edge case
                                ' documented in this file's own class remarks far less likely to matter in
                                ' practice, since every tracking number starts genuinely absent for this
                                ' billing date at the start of each run.
                                Try
                                    Using deleteCmd As New SqlCommand("DELETE FROM FedEx WHERE Billing_Start_Date = @BillStart", conn)
                                        deleteCmd.Parameters.AddWithValue("@BillStart", billStartDate)
                                        deleteCmd.ExecuteNonQuery()
                                    End Using
                                Catch ex As Exception
                                    ErrorLogHelper.LogError("Upload FedEx Charges", $"SQL error in: DELETE FROM FedEx - {ex.Message}")
                                    errorCount += 1
                                End Try

                                Using reader = New StreamReader(csvFilePath)
                                    Using csv = New CsvReader(reader, CultureInfo.InvariantCulture)
                                        csv.Read() ' header row
                                        csv.ReadHeader()

                                        While csv.Read()
                                            Dim accountNum = GetField(csv, 2)
                                            If String.IsNullOrEmpty(accountNum) Then Exit While ' matches "While ws.Cells(i, 2) <> """

                                            If accountNum <> TargetAccountNumber Then Continue While

                                            Try
                                                ProcessRow(conn, csv, billStartDate)
                                            Catch ex As SqlException When IsDuplicateKeyViolation(ex)
                                                ' Matches the original's specific suppression of Access error
                                                ' 3022 (duplicate key violation) - silently skipped, NOT
                                                ' logged, matching the original's own deliberate choice here.
                                            Catch ex As Exception
                                                ErrorLogHelper.LogError("Upload FedEx Charges", $"SQL error: {ex.Message}")
                                                errorCount += 1
                                            End Try
                                        End While
                                    End Using
                                End Using
                            End Using

                            Return errorCount
                        End Function)
    End Function

    Private Function IsDuplicateKeyViolation(ex As SqlException) As Boolean
        Return ex.Number = 2627 OrElse ex.Number = 2601
    End Function

    Private Sub ProcessRow(conn As SqlConnection, csv As CsvReader, billStartDate As Date)
        Dim rawDate = GetField(csv, 3) ' YYYYMMDD, reassembled below matching the original's own Mid/Right/Left slicing
        Dim tranDate = New Date(CInt(rawDate.Substring(0, 4)), CInt(rawDate.Substring(4, 2)), CInt(rawDate.Substring(6, 2)))
        Dim trackingNum = GetField(csv, 10)
        Dim newCost = ParseCostOrZero(GetField(csv, 12))

        Dim existingCost = LookupExistingTotalCost(conn, trackingNum, billStartDate)

        If existingCost = 0 Then
            Using cmd As New SqlCommand("INSERT INTO FedEx (Tracking_Num, Transaction_Date, Total_Cost, Billing_Start_Date) VALUES (@TrackingNum, @TranDate, @Cost, @BillStart)", conn)
                cmd.Parameters.AddWithValue("@TrackingNum", trackingNum)
                cmd.Parameters.AddWithValue("@TranDate", tranDate)
                cmd.Parameters.AddWithValue("@Cost", newCost)
                cmd.Parameters.AddWithValue("@BillStart", billStartDate)
                cmd.ExecuteNonQuery()
            End Using
        Else
            Dim updatedCost = existingCost + newCost
            Using cmd As New SqlCommand("UPDATE FedEx SET Total_Cost = @Cost WHERE Tracking_Num = @TrackingNum AND Billing_Start_Date = @BillStart", conn)
                cmd.Parameters.AddWithValue("@Cost", updatedCost)
                cmd.Parameters.AddWithValue("@TrackingNum", trackingNum)
                cmd.Parameters.AddWithValue("@BillStart", billStartDate)
                cmd.ExecuteNonQuery()
            End Using
        End If
    End Sub

    Private Function LookupExistingTotalCost(conn As SqlConnection, trackingNum As String, billStartDate As Date) As Decimal
        Using cmd As New SqlCommand("SELECT Total_Cost FROM FedEx WHERE Tracking_Num = @TrackingNum AND Billing_Start_Date = @BillStart", conn)
            cmd.Parameters.AddWithValue("@TrackingNum", trackingNum)
            cmd.Parameters.AddWithValue("@BillStart", billStartDate)
            Dim result = cmd.ExecuteScalar()
            If result Is Nothing OrElse result Is DBNull.Value Then Return 0
            Return Convert.ToDecimal(result)
        End Using
    End Function

    Private Function GetField(csv As CsvReader, oneIndexedColumn As Integer) As String
        Return If(csv.GetField(oneIndexedColumn - 1), String.Empty).Trim()
    End Function

    Private Function ParseCostOrZero(s As String) As Decimal
        Dim result As Decimal
        Decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, result)
        Return result
    End Function

End Module
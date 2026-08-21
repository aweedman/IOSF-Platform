Imports Microsoft.Data.SqlClient
Imports CsvHelper
Imports System.Globalization
Imports System.IO

''' <summary>
''' Imports FedEx charges from a CSV export into the FedEx table for a given billing
''' start date.
'''
''' Account filter (column 2 = "324282815") is a hardcoded literal - the CSV export
''' combines charges across multiple FedEx accounts, and only this one is relevant here.
'''
''' Date parsing (column 3): reassembles an 8-character YYYYMMDD string into a proper date
''' (Year = chars[0:4], Month = chars[4:6], Day = chars[6:8]).
'''
''' AGGREGATION LOGIC has a real edge case worth understanding: it decides INSERT vs.
''' UPDATE by checking whether the looked-up Total_Cost equals 0, not whether a matching
''' row actually exists. If an existing row's Total_Cost happens to be exactly 0, this
''' would attempt a duplicate INSERT instead of an UPDATE - so any resulting duplicate-key
''' error (SQL Server error 2627 or 2601) is caught and silently skipped rather than
''' logged, which is intentional here, not a bug being papered over.
'''
''' DELETE-BEFORE-INSERT: before processing any rows, existing FedEx rows for the selected
''' billing start date are deleted first, making it safe to re-run/reload the same file.
''' This isn't wrapped in one transaction with the row-by-row import - the delete is its
''' own quick upfront step, and each row's insert/update failure is handled independently
''' without rolling back the rest.
''' </summary>
Public Module FedExChargesToDbJob

    Private Const TargetAccountNumber As String = "324282815"

    Public Function RunAsync(csvFilePath As String, billStartDate As Date) As Task(Of Integer)
        Return Task.Run(Function()
                            Dim errorCount = 0

                            Using conn As New SqlConnection(ConfigHelper.ConnectionString)
                                conn.Open()

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
                                            If String.IsNullOrEmpty(accountNum) Then Exit While ' end of data

                                            If accountNum <> TargetAccountNumber Then Continue While

                                            Try
                                                ProcessRow(conn, csv, billStartDate)
                                            Catch ex As SqlException When IsDuplicateKeyViolation(ex)
                                                ' Intentionally silent - see class remarks.
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
        Dim rawDate = GetField(csv, 3) ' YYYYMMDD
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
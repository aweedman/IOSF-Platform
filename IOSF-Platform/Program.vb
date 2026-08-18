Imports System.Windows.Forms

''' <summary>
''' Application entry point. Replaces the old "msaccess.exe /x FunctionName" pattern:
''' Task Scheduler now calls "IOSF-Platform.exe JobName" for headless runs, or the .exe
''' with no arguments to launch the interactive UI.
'''
''' SETUP REQUIRED IN VISUAL STUDIO: this only runs if the project's Startup object is
''' set to Sub Main instead of the default Form1. Project menu -> [ProjectName] Properties
''' -> Application tab -> change "Startup object" to "Sub Main". Also uncheck "Enable
''' application framework" if it's checked, since that framework expects to own startup
''' itself and will conflict with a custom Sub Main.
'''
''' JOB NAMES below match the original scheduled macro file names exactly (verified
''' against macros/*.bas, not guessed) - AfterHours, EarlyMeeting, RemoteLock,
''' SpheremailStorageStaff, SpheremailStorageCustomers, CustomerMaster, CallCounts,
''' CustomerXref, IncomeDB, and PaperCutFrequentScanners are fully wired.
''' AuthorizeDotNet and TomorrowRoomUsage throw a distinct error - their source functions
''' (AuthNet, DailyRoomUsage) were not found anywhere in the exported repository at all,
''' so those need to be located/exported from the live Access app before they can be
''' ported, independent of anything else here.
''' </summary>
Module Program

    <STAThread()>
    Sub Main(args As String())
        ' REQUIRED for QODBC: .NET Core/.NET 5+ (unlike .NET Framework) doesn't include
        ' legacy code pages like Windows-1252 by default - QODBC's driver relies on this
        ' encoding for text data and throws System.NotSupportedException ("No data is
        ' available for encoding 1252...") on any query touching text columns without it.
        ' Must run once, before ANY QODBC operation - confirmed via a real crash in
        ' KubeInvoicesToQbJob's very first ODBC query. This isn't specific to that one
        ' job; it would affect every other QODBC-based job in this app the same way,
        ' just hadn't been hit yet by whatever query ran first in each of those.
        ' REQUIRES ADDING THE System.Text.Encoding.CodePages NUGET PACKAGE to the project
        ' (same compile-test caveat as Ical.Net/ClosedXML elsewhere in this port).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)

        AppConfig.Load()
        QuestPDF.Settings.License = QuestPDF.Infrastructure.LicenseType.Community

        If args.Length = 0 Then
            RunInteractive()
        Else
            RunHeadlessJob(args(0)).GetAwaiter().GetResult()
        End If
    End Sub

    Private Sub RunInteractive()
        Application.SetHighDpiMode(HighDpiMode.SystemAware)
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        ' LandingPageForm is the real dashboard, covering every job ported so far - see its
        ' own doc comment. MainForm.vb (frmMain.cls's port - Exit button + DateTimePicker)
        ' still exists as a file but is no longer the interactive entry point; it had no
        ' real logic of its own and isn't needed as a separate launch target.
        Application.Run(New LandingPageForm())
    End Sub

    ''' <summary>
    ''' Runs one named job headlessly and reports the result the same way the original's
    ''' "Batch <> X" branch did: silence on success, an EmailError on failure - nobody's
    ''' watching a scheduled task's console output, so a MsgBox (used for interactive runs)
    ''' would never be seen here.
    ''' </summary>
    Private Async Function RunHeadlessJob(jobName As String) As Task
        Try
            Select Case jobName
                Case "EarlyMeeting"
                    ' Self-reports via EmailHelper.EmailError internally - original had no
                    ' Batch <> "X" check at all, always headless-style.
                    Await EarlyMeetingJob.RunAsync()

                Case "AfterHours"
                    ' Same as above - self-reporting, no Batch check in the original.
                    Await AfterHoursJob.RunAsync(daysBack:=1)

                Case "SpheremailStorageStaff"
                    Dim errorCount = Await SphereMailStorageEmailJob.RunAsync(SphereMailStorageEmailJob.Mode.StaffSummary)
                    ReportBatchResult("Spheremail Storage Emails (Staff)", errorCount)

                Case "SpheremailStorageCustomers"
                    Dim errorCount = Await SphereMailStorageEmailJob.RunAsync(SphereMailStorageEmailJob.Mode.IndividualCustomers)
                    ReportBatchResult("Spheremail Storage Emails (Customers)", errorCount)

                Case "RemoteLock"
                    Dim errorCount = Await RemoteLockUsersJob.RunAsync()
                    ReportBatchResult("RemoteLock Users", errorCount)

                Case "CustomerXref"
                    Dim errors = Await CustomerXrefJob.RunAsync()
                    If errors > 0 Then EmailHelper.EmailError($"Update Customer XRef: {errors} errors in log.")

                Case "CustomerMaster"
                    Dim errors = Await CustomerMasterDeltaJob.RunAsync()
                    If errors > 0 Then EmailHelper.EmailError($"Update Customer Master: {errors} errors in log.")

                Case "PaperCutFrequentScanners"
                    Dim errors = PaperCutSyncJob.Run()
                    If errors > 0 Then EmailHelper.EmailError($"Papercut Scan Actions: {errors} errors in log.")

                Case "IncomeDB"
                    Dim monthStart = New Date(Date.Today.Year, Date.Today.Month, 1)
                    Dim monthEnd = monthStart.AddMonths(1).AddDays(-1)
                    Dim errors = Await IncomeDbJob.RunAsync(monthStart, monthEnd)
                    If errors > 0 Then EmailHelper.EmailError($"Upload Income: {errors} errors in log.")

                Case "CallCounts"
                    Dim startDate = CallCountsJob.GetNextStartDate()
                    Dim endDate = Date.Today.AddDays(-1)
                    If startDate > endDate Then
                        ' Matches original: nothing new to process, do nothing (not an error).
                    Else
                        Dim errors = Await CallCountsJob.RunAsync(startDate, endDate)
                        If errors > 0 Then EmailHelper.EmailError($"Call Counts to DB: {errors} errors in log.")
                    End If

                Case "AuthorizeDotNet"
                    Throw New NotImplementedException(
                        "AuthorizeDotNet calls AuthNet(), which was NOT FOUND anywhere in the exported repository " &
                        "source - this isn't just unported, the source itself is missing. Check the live Access " &
                        "application directly and export it via the VCS add-in before this can be ported.")

                Case "TomorrowRoomUsage"
                    Throw New NotImplementedException(
                        "TomorrowRoomUsage calls DailyRoomUsage(), which was NOT FOUND anywhere in the exported " &
                        "repository source - this isn't just unported, the source itself is missing. Check the live " &
                        "Access application directly and export it via the VCS add-in before this can be ported.")

                Case Else
                    Throw New ArgumentException($"Unknown job name: '{jobName}'")
            End Select

        Catch ex As Exception
            ' Catches anything a job's own error handling didn't - e.g. AppConfig.Load
            ' succeeding but the job itself throwing something genuinely unexpected.
            ErrorLogHelper.LogError(jobName, $"Unhandled error running job '{jobName}': {ex.Message}")
            EmailHelper.EmailError($"{jobName}: unhandled error - {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' Mirrors the original's "ElseIf Batch = 'X' And RoutineErrors &lt;&gt; 0 Then
    ''' emailerror(...)" branch - silence on success, since nobody's watching a scheduled
    ''' task run. The interactive equivalent (MsgBox "Process Complete. N errors.") belongs
    ''' in the actual button click handler once the UI is built - same error count, just a
    ''' different way of surfacing it depending on how the job was triggered.
    ''' </summary>
    Private Sub ReportBatchResult(jobLabel As String, errorCount As Integer)
        If errorCount > 0 Then
            EmailHelper.EmailError($"{jobLabel}: {errorCount} errors in log.")
        End If
    End Sub

End Module
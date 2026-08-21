Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Interactive dashboard for running every background job in this app manually - each
''' section groups related jobs (Quickbooks Interfaces / SQL Server Interfaces / Evo
''' Interfaces / Reports / Other / Tables), and each button runs one job end to end
''' (prompting for whatever inputs it needs, running it, and logging the result).
'''
''' Some buttons are shown disabled with a "Not yet ported" tooltip - these represent
''' features that don't have a working implementation behind them yet, so the dashboard
''' also serves as a visual checklist of what's built vs. still pending. Several button
''' labels carry a leading number, which lines up with this team's own external task
''' tracker for cross-reference.
'''
''' LAYOUT: the outer scrollable container is a plain Panel with manual vertical
''' positioning (tracked via currentY), not a FlowLayoutPanel - nesting a wrap-enabled
''' FlowLayoutPanel inside another FlowLayoutPanel's own layout engine produces
''' unpredictable width measurement (buttons flow off-screen horizontally instead of
''' wrapping, with no way to reach later sections). A plain outer Panel with explicit
''' Y-tracking avoids that nested-layout ambiguity entirely. Each section's button grid
''' still uses a FlowLayoutPanel (WrapContents=True) internally, but with its width
''' pinned via matching MinimumSize/MaximumSize (both set to the same width, height
''' uncapped) - the standard, reliable WinForms way to get "fixed width, auto height,
''' wraps" instead of just growing horizontally forever.
'''
''' To wire up a still-disabled placeholder button once its job is implemented: find it
''' in BuildLayout, swap AddPlaceholder for AddButton with a handler, and add that handler
''' below following the pattern of the existing working ones - just needs a call into
''' RunJobAsync/RunSelfReportingJobAsync; no layout change needed.
'''
''' Interactive date/mode selection (DateRangeDialog, the SphereMail Yes/No prompt, the
''' RemoteLock auth dialog) lives here rather than in the job classes themselves - job
''' logic takes explicit parameters; this form decides how to obtain them.
''' </summary>
Public Class LandingPageForm
    Inherits Form

    Private Const SectionWidth As Integer = 560
    Private Const LeftMargin As Integer = 10

    Private outerPanel As Panel
    Private logBox As TextBox
    Private jobButtons As New List(Of Button) ' only REAL (enabled) buttons - see SetButtonsEnabled
    Private currentY As Integer = 10

    Public Sub New()
        Text = "IOSF Platform - Job Dashboard"
        Width = 1050
        Height = 700
        StartPosition = FormStartPosition.CenterScreen

        outerPanel = New Panel With {
            .Dock = DockStyle.Left,
            .Width = 600,
            .AutoScroll = True
        }

        logBox = New TextBox With {
            .Dock = DockStyle.Fill,
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Vertical,
            .Font = New Font("Consolas", 9)
        }

        Controls.Add(logBox)
        Controls.Add(outerPanel)

        BuildLayout()

        AppendLog("Ready.")
    End Sub

    ''' <summary>Builds every section and its buttons. See class remarks for the layout approach.</summary>
    Private Sub BuildLayout()
        Dim qb = AddSection("Quickbooks Interfaces")
        AddButton(qb, "Kube Invoices to QB", AddressOf RunKubeInvoicesToQb)
        AddButton(qb, "Kube Payment to QB", AddressOf RunKubePaymentsToQb)
        FinishSection(qb)

        Dim sql = AddSection("SQL Server Interfaces")
        AddButton(sql, "150.1 - Variable Charges to DB", AddressOf RunVariableCharges)
        AddButton(sql, "160.1 - SendPro Forwards to DB", AddressOf RunSendProForwards)
        AddButton(sql, "180.1 - Spheremail Charges to DB", AddressOf RunSpheremailCharges)
        AddButton(sql, "160.2 FedEx Charges to DB", AddressOf RunFedExCharges)
        AddButton(sql, "160.3 Edit SendPro", AddressOf RunEditSendPro)
        AddButton(sql, "190.1 Kube Meetings to DB", AddressOf RunKubeMeetings)
        AddButton(sql, "QB Customer Master to DB - Delta", AddressOf RunCustomerMaster)
        AddButton(sql, "Income to DB...", AddressOf RunIncomeDb)
        AddButton(sql, "190.3 - Call Counts to DB...", AddressOf RunCallCounts)
        AddButton(sql, "QB Customer Master to DB - Full", AddressOf RunCustomerMasterFull)
        AddButton(sql, "PnL to DB", AddressOf RunPnLToDb)
        AddButton(sql, "Evo Customer XRef to DB", AddressOf RunCustomerXref)
        FinishSection(sql)

        Dim evo = AddSection("Evo Interfaces")
        AddButton(evo, "140.2 Copier to Evo", AddressOf RunCopierChargesToEvo)
        AddButton(evo, "150.2 Scan Extra Pages to Evo", AddressOf RunScanExtraPagesToEvo)
        AddButton(evo, "160.4 Forwards to Evo", AddressOf RunMailForwardsToEvo)
        AddButton(evo, "180.2 SphereMail to Evo", AddressOf RunSpheremailToEvo)
        FinishSection(evo)

        Dim reports = AddSection("Reports")
        AddButton(reports, "140.1 Copier Counts", AddressOf RunCopierCountsReport)
        AddButton(reports, "180.3 - Spheremail Storage Report", AddressOf RunSpheremailStorageReport)
        AddButton(reports, "190.2 - Room Usage Report", AddressOf RunRoomUsageReport)
        AddButton(reports, "190.4 Call Counts", AddressOf RunCallCountsReport)
        AddButton(reports, "Class Checks", AddressOf RunClassChecks)
        AddButton(reports, "Mail Forwards", AddressOf RunMailForwardsReport)
        AddPlaceholder(reports, "IA Revenue per Customer")
        FinishSection(reports)

        Dim other = AddSection("Other")
        AddButton(other, "RemoteLock Users", AddressOf RunRemoteLockUsers)
        AddButton(other, "Spheremail Storage Emails...", AddressOf RunSpheremailStorage)
        AddButton(other, "Afterhours Room Usage Emails", AddressOf RunAfterHours)
        AddButton(other, "Early Morning Meeting Emails", AddressOf RunEarlyMeeting)
        AddButton(other, "RemoteLock Refresh Token...", AddressOf RunRemoteLockAuth)
        AddButton(other, "Spheremail Worklist", AddressOf RunSpheremailWorklist)
        AddButton(other, "Papercut Scan Actions and Users", AddressOf RunPaperCut)
        AddButton(other, "Edit Customer Master", AddressOf RunCustomerMasterEditor)
        AddButton(other, "Random Facility Code", AddressOf RunRandomFacilityCode)
        FinishSection(other)

        ' Direct table view/add/edit/delete access for these specific tables.
        ' IO_Employees is read-only (it's actually a view, not a table).
        Dim tables = AddSection("Tables")
        AddButton(tables, "Answering_Config", AddressOf RunEditAnsweringConfig)
        AddButton(tables, "Config", AddressOf RunEditConfig)
        AddButton(tables, "Error_Log", AddressOf RunEditErrorLog)
        AddButton(tables, "Holidays", AddressOf RunEditHolidays)
        AddButton(tables, "IO_Employees (read-only)", AddressOf RunEditIoEmployees)
        AddButton(tables, "SendPro_XRef", AddressOf RunEditSendProXref)
        FinishSection(tables)
    End Sub

    ''' <summary>
    ''' Adds a bold section header + divider line to outerPanel at the current Y position,
    ''' advances currentY past them, and returns a new wrap-enabled FlowLayoutPanel
    ''' (positioned at the new currentY, width PINNED to SectionWidth) for that section's
    ''' buttons. Caller must call FinishSection() once all buttons for this section have
    ''' been added, to advance currentY past this section's actual (wrapped) height before
    ''' the next section starts.
    ''' </summary>
    Private Function AddSection(title As String) As FlowLayoutPanel
        Dim header As New Label With {
            .Text = title,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(LeftMargin, currentY)
        }
        outerPanel.Controls.Add(header)
        currentY += header.PreferredHeight + 4

        Dim divider As New Panel With {
            .Location = New Point(LeftMargin, currentY),
            .Size = New Size(SectionWidth, 2),
            .BackColor = Color.SteelBlue
        }
        outerPanel.Controls.Add(divider)
        currentY += divider.Height + 8

        Dim sectionPanel As New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Location = New Point(LeftMargin, currentY),
            .MinimumSize = New Size(SectionWidth, 0),
            .MaximumSize = New Size(SectionWidth, 0) ' pinning Min=Max=SectionWidth locks the WIDTH exactly, while height (0 = uncapped in both) grows freely to fit however many rows wrapping produces
        }
        outerPanel.Controls.Add(sectionPanel)

        Return sectionPanel
    End Function

    ''' <summary>
    ''' Advances currentY past this section's actual height (now that all its buttons have
    ''' been added and its wrapped height is known), ready for the next section to start
    ''' below it rather than overlapping it.
    ''' </summary>
    Private Sub FinishSection(sectionPanel As FlowLayoutPanel)
        sectionPanel.PerformLayout() ' force a fresh layout pass before reading .Bottom, rather than assuming AutoSize already recalculated
        currentY = sectionPanel.Bottom + 10
    End Sub

    ''' <summary>A real, working button - wired to an actual handler and enabled.</summary>
    Private Sub AddButton(section As FlowLayoutPanel, label As String, handler As EventHandler)
        Dim btn As New Button With {.Text = label, .Width = 175, .Height = 48, .Margin = New Padding(0, 0, 8, 8)}
        AddHandler btn.Click, handler
        section.Controls.Add(btn)
        jobButtons.Add(btn)
    End Sub

    ''' <summary>
    ''' A not-yet-ported button from the original Landing Page - shown disabled with a
    ''' tooltip, so the dashboard stays a complete, honest checklist against the original
    ''' (and against Asana, for buttons whose label includes a step number).
    ''' </summary>
    Private Sub AddPlaceholder(section As FlowLayoutPanel, label As String)
        Dim btn As New Button With {.Text = label, .Width = 175, .Height = 48, .Margin = New Padding(0, 0, 8, 8), .Enabled = False}
        Dim tip As New ToolTip()
        tip.SetToolTip(btn, "Not yet ported")
        section.Controls.Add(btn)
        ' Deliberately NOT added to jobButtons - it's permanently disabled, no need for
        ' SetButtonsEnabled to touch it.
    End Sub

    Private Sub AppendLog(text As String)
        logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {text}{Environment.NewLine}")
        ' AppendText alone doesn't reliably scroll the view to follow newly-appended text
        ' (especially when called repeatedly via BeginInvoke from a background thread, as
        ' Kube Invoices/Payments' per-item logging does) - a job with enough log lines to
        ' overflow the visible area could leave its final "completed" message scrolled out
        ' of view, even though it was genuinely written. Explicitly moving the caret to
        ' the end and scrolling to it fixes this for every job using AppendLog.
        logBox.SelectionStart = logBox.Text.Length
        logBox.ScrollToCaret()
    End Sub

    Private Sub SetButtonsEnabled(enabled As Boolean)
        For Each btn In jobButtons
            btn.Enabled = enabled
        Next
    End Sub

    ''' <summary>
    ''' Shared wrapper for jobs that return an error count. Disables buttons while running,
    ''' logs start/completion/exception, re-enables afterward - avoids repeating this in
    ''' every single handler below.
    ''' </summary>
    Private Async Function RunJobAsync(name As String, job As Func(Of Task(Of Integer))) As Task
        SetButtonsEnabled(False)
        AppendLog($"Running {name}...")
        Try
            Dim errors = Await job()
            If errors = 0 Then
                AppendLog($"{name}: completed successfully.")
            Else
                AppendLog($"{name}: completed with {errors} error(s) - see Error_Log.")
            End If
        Catch ex As Exception
            AppendLog($"{name}: FAILED - {ex.Message}")
        Finally
            SetButtonsEnabled(True)
        End Try
    End Function

    ''' <summary>
    ''' Same as RunJobAsync, for jobs that don't return a count (they email/log their own
    ''' errors internally - EarlyMeetingJob, AfterHoursJob). The default success message
    ''' hedges ("check email/Error_Log for any issues") because those two jobs' Catch
    ''' blocks swallow failures internally (log/email, no rethrow) - reaching this line
    ''' does NOT guarantee nothing went wrong for them. RemoteLock Refresh Token is
    ''' different: ExchangeAuthorizationCodeAsync throws via EnsureSuccess() on any
    ''' failure, so reaching this line for that job means it genuinely, unambiguously
    ''' succeeded - the hedge was misleading there, so its call site now passes its own
    ''' accurate message instead of using this default.
    ''' </summary>
    Private Async Function RunSelfReportingJobAsync(name As String, job As Func(Of Task),
                                                      Optional successMessage As String = Nothing) As Task
        SetButtonsEnabled(False)
        AppendLog($"Running {name}...")
        Try
            Await job()
            AppendLog($"{name}: {If(successMessage, "completed (check email/Error_Log for any issues).")}")
        Catch ex As Exception
            AppendLog($"{name}: FAILED - {ex.Message}")
        Finally
            SetButtonsEnabled(True)
        End Try
    End Function

    Private Async Sub RunKubeMeetings(sender As Object, e As EventArgs)
        Using dlg As New OpenFileDialog With {
            .Title = "Select the Excel file to process",
            .Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb",
            .Multiselect = False
        }
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunJobAsync("Kube Meetings to DB", Function() KubeMeetingsToDbJob.RunAsync(dlg.FileName))
        End Using
    End Sub

    Private Sub RunClassChecks(sender As Object, e As EventArgs)
        Using typeDlg As New ClassCheckTypeDialog()
            If typeDlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim selectedType = typeDlg.SelectedType
            If selectedType = 0 Then Return ' user cancelled the type selection

            Dim defaultDate = DefaultDateHelper.ComputeDefaultDate(1, -1) ' 1st of last month
            Using dateDlg As New SingleDateDialog("Class Checks", "From Date", defaultDate)
                If dateDlg.ShowDialog(Me) <> DialogResult.OK Then Return

                Try
                    Cursor = Cursors.WaitCursor
                    Dim table = ClassChecksJob.RunCheck(selectedType, dateDlg.SelectedDate)
                    Dim label = ClassChecksJob.TypeLabels(selectedType)

                    Dim grid As New DataGridView With {
                        .Dock = DockStyle.Fill,
                        .ReadOnly = True,
                        .AllowUserToAddRows = False,
                        .AllowUserToDeleteRows = False,
                        .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                        .SelectionMode = DataGridViewSelectionMode.CellSelect,
                        .DataSource = table
                    }
                    AddHandler grid.DataError, Sub(s, ev) ev.ThrowException = False

                    Using resultForm As New Form With {.Text = $"Class Checks - {label}", .Width = 900, .Height = 600, .StartPosition = FormStartPosition.CenterScreen}
                        resultForm.Controls.Add(grid)
                        resultForm.ShowDialog(Me)
                    End Using
                Catch ex As Exception
                    MessageBox.Show(Me, $"Process Abended: {ex.Message}", "Class Checks", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    Cursor = Cursors.Default
                End Try
            End Using
        End Using
    End Sub

    Private Async Sub RunSpheremailStorageReport(sender As Object, e As EventArgs)
        Dim defaultDate = DefaultDateHelper.ComputeDefaultDate(1, 1) ' 1st of NEXT month
        Using dlg As New SingleDateDialog("Spheremail Storage Report", "Invoice Date", defaultDate)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return

            Try
                Cursor = Cursors.WaitCursor
                Dim days = SpheremailStorageReportHelper.ComputeDaysFromInvoiceDate(dlg.SelectedDate)
                Dim result = Await SphereMailStorageJob.RunAsync(days)

                ' Filtered to San Francisco, matching SphereMailStorageEmailJob's own
                ' filter for its staff-summary PDF - same location this whole port has
                ' treated as the only active one (Burlingame decommissioned).
                Const location = "San Francisco"
                Dim locationRows = result.Rows.Where(Function(r) r.Location = location).ToList()

                If locationRows.Count = 0 Then
                    MessageBox.Show(Me, "No storage items found for this date.", "Spheremail Storage Report")
                    Return
                End If

                ' Reuses the same PDF generator used for the Spheremail Storage email
                ' attachments - not a separate grid/print implementation.
                Dim pdfPath = IO.Path.Combine(IO.Path.GetTempPath(), $"Spheremail Storage Report {DateTime.Now:yyyyMMdd_HHmmss}.pdf")
                Await ReportGenerator.GenerateSphereMailStoragePdfAsync(locationRows, location, pdfPath)

                ' Opens directly with the system's default PDF viewer - no intermediate grid step.
                Process.Start(New Diagnostics.ProcessStartInfo(pdfPath) With {.UseShellExecute = True})

                If result.ErrorCount > 0 Then
                    MessageBox.Show(Me, $"Report opened. {result.ErrorCount} error(s) occurred fetching data - see Error_Log.", "Spheremail Storage Report")
                End If
            Catch ex As Exception
                MessageBox.Show(Me, $"Error generating report: {ex.Message}", "Spheremail Storage Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Cursor = Cursors.Default
            End Try
        End Using
    End Sub

    Private Async Sub RunSpheremailWorklist(sender As Object, e As EventArgs)
        Try
            Cursor = Cursors.WaitCursor
            Dim rows = Await SpheremailWorklistJob.FetchWorklist()

            If rows.Count = 0 Then
                MessageBox.Show(Me, "No worklist items found.", "Spheremail Worklist")
                Return
            End If

            ' Reuses ReportGenerator, same pattern as Spheremail Storage Report - no
            ' intermediate grid window.
            Dim pdfPath = IO.Path.Combine(IO.Path.GetTempPath(), $"Spheremail Worklist {DateTime.Now:yyyyMMdd_HHmmss}.pdf")
            Await ReportGenerator.GenerateSphereMailWorklistPdfAsync(rows, pdfPath)

            Process.Start(New Diagnostics.ProcessStartInfo(pdfPath) With {.UseShellExecute = True})
        Catch ex As Exception
            MessageBox.Show(Me, $"Error generating worklist: {ex.Message}", "Spheremail Worklist", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub RunMailForwardsReport(sender As Object, e As EventArgs)
        Dim defaultDate = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Using dlg As New SingleDateDialog("Mail Forwards Report", "Billing Cycle Start Date", defaultDate)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Using reportForm As New MailForwardsReportForm(dlg.SelectedDate)
                reportForm.ShowDialog(Me)
            End Using
        End Using
    End Sub

    Private Sub RunCallCountsReport(sender As Object, e As EventArgs)
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using dlg As New DateRangeDialog("Call Counts Report", "Bill From Date", "Bill To Date", defaultFrom, defaultTo)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Using reportForm As New CallCountsReportForm(dlg.FromDate, dlg.ToDate)
                reportForm.ShowDialog(Me)
            End Using
        End Using
    End Sub

    Private Sub RunCopierCountsReport(sender As Object, e As EventArgs)
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using dlg As New DateRangeDialog("Copier Counts Report", "Bill From Date", "Bill To Date", defaultFrom, defaultTo)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Using reportForm As New CopierCountsReportForm(dlg.FromDate, dlg.ToDate)
                reportForm.ShowDialog(Me)
            End Using
        End Using
    End Sub

    Private Async Sub RunSpheremailToEvo(sender As Object, e As EventArgs)
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using rangeDlg As New DateRangeDialog("Spheremail to Evo", "Billing Cycle Start Date", "Billing Cycle End Date", defaultFrom, defaultTo)
            If rangeDlg.ShowDialog(Me) <> DialogResult.OK Then Return

            Using dateDlg As New SingleDateDialog("Spheremail to Evo", "Posting Date", defaultTo)
                If dateDlg.ShowDialog(Me) <> DialogResult.OK Then Return
                Await RunJobAsync("Spheremail to Evo", Function() SpheremailToEvoJob.RunAsync(rangeDlg.FromDate, rangeDlg.ToDate, dateDlg.SelectedDate))
            End Using
        End Using
    End Sub

    Private Async Sub RunMailForwardsToEvo(sender As Object, e As EventArgs)
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using rangeDlg As New DateRangeDialog("Mail Forwards to Evo", "Billing Cycle Start Date", "Billing Cycle End Date", defaultFrom, defaultTo)
            If rangeDlg.ShowDialog(Me) <> DialogResult.OK Then Return

            Using dateDlg As New SingleDateDialog("Mail Forwards to Evo", "Posting Date", defaultTo)
                If dateDlg.ShowDialog(Me) <> DialogResult.OK Then Return
                Await RunJobAsync("Mail Forwards to Evo", Function() MailForwardsToEvoJob.RunAsync(rangeDlg.FromDate, rangeDlg.ToDate, dateDlg.SelectedDate))
            End Using
        End Using
    End Sub

    Private Async Sub RunScanExtraPagesToEvo(sender As Object, e As EventArgs)
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using rangeDlg As New DateRangeDialog("Scan Extra Pages to Evo", "Billing Cycle Start Date", "Billing Cycle End Date", defaultFrom, defaultTo)
            If rangeDlg.ShowDialog(Me) <> DialogResult.OK Then Return

            Using dateDlg As New SingleDateDialog("Scan Extra Pages to Evo", "Posting Date", defaultTo)
                If dateDlg.ShowDialog(Me) <> DialogResult.OK Then Return
                Await RunJobAsync("Scan Extra Pages to Evo", Function() ScanExtraPagesToEvoJob.RunAsync(rangeDlg.FromDate, rangeDlg.ToDate, dateDlg.SelectedDate))
            End Using
        End Using
    End Sub

    Private Async Sub RunCopierChargesToEvo(sender As Object, e As EventArgs)
        ' Three dates: billing cycle start/end (same 26th-of-last-month through
        ' 25th-of-this-month defaults as other billing-cycle jobs), plus a separate
        ' Posting Date (defaults to the 25th of THIS month, i.e. "today's billing
        ' cycle" rather than last month's).
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using rangeDlg As New DateRangeDialog("Copier Charges to Evo", "Billing Cycle Start Date", "Billing Cycle End Date", defaultFrom, defaultTo)
            If rangeDlg.ShowDialog(Me) <> DialogResult.OK Then Return

            Using dateDlg As New SingleDateDialog("Copier Charges to Evo", "Posting Date", defaultTo)
                If dateDlg.ShowDialog(Me) <> DialogResult.OK Then Return
                Await RunJobAsync("Copier Charges to Evo", Function() CopierChargesToEvoJob.RunAsync(rangeDlg.FromDate, rangeDlg.ToDate, dateDlg.SelectedDate))
            End Using
        End Using
    End Sub

    Private Async Sub RunFedExCharges(sender As Object, e As EventArgs)
        Dim defaultDate = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Using dateDlg As New SingleDateDialog("FedEx Charges to DB", "Billing Cycle Start Date", defaultDate)
            If dateDlg.ShowDialog(Me) <> DialogResult.OK Then Return

            Using fileDlg As New OpenFileDialog With {
                .Title = "Select the Report",
                .Filter = "CSV Files|*.csv",
                .Multiselect = False
            }
                If fileDlg.ShowDialog(Me) <> DialogResult.OK Then Return
                Await RunJobAsync("FedEx Charges to DB", Function() FedExChargesToDbJob.RunAsync(fileDlg.FileName, dateDlg.SelectedDate))
            End Using
        End Using
    End Sub

    Private Async Sub RunSendProForwards(sender As Object, e As EventArgs)
        Using dlg As New OpenFileDialog With {
            .Title = "Select the Report",
            .Filter = "CSV Files|*.csv",
            .Multiselect = False
        }
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunJobAsync("SendPro Forwards to DB", Function() SendProForwardsToDbJob.RunAsync(dlg.FileName))
        End Using
    End Sub

    Private Async Sub RunKubeInvoicesToQb(sender As Object, e As EventArgs)
        Using dlg As New OpenFileDialog With {
            .Title = "Select the Excel file to process",
            .Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb",
            .Multiselect = False
        }
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            ' Logs each created Invoice/Credit Memo to the UI log as it happens. The job
            ' runs on a background thread, so this callback marshals back to the UI
            ' thread. Uses Invoke (blocking), NOT BeginInvoke (queued/async): with
            ' BeginInvoke, the background job could finish and return before an
            ' earlier-queued per-item message had actually been appended, letting
            ' RunJobAsync's own "completed successfully" message (appended directly on
            ' the UI thread once the job returns) land ahead of it in the log. Invoke
            ' blocks the background thread until AppendLog has actually run, guaranteeing
            ' every per-item message is written before the job can proceed/return.
            Dim logCallback As Action(Of String) = Sub(msg) Invoke(New Action(Sub() AppendLog(msg)))
            Await RunJobAsync("Kube Invoices to QB", Function() KubeInvoicesToQbJob.RunAsync(dlg.FileName, logCallback))
        End Using
    End Sub

    Private Async Sub RunKubePaymentsToQb(sender As Object, e As EventArgs)
        Using dlg As New OpenFileDialog With {
            .Title = "Select the Excel file to process",
            .Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb",
            .Multiselect = False
        }
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            ' Same fix as RunKubeInvoicesToQb above - Invoke, not BeginInvoke, for the
            ' same ordering-guarantee reason.
            Dim logCallback As Action(Of String) = Sub(msg) Invoke(New Action(Sub() AppendLog(msg)))
            Await RunJobAsync("Kube Payments to QB", Function() KubePaymentsToQbJob.RunAsync(dlg.FileName, logCallback))
        End Using
    End Sub

    Private Async Sub RunCustomerXref(sender As Object, e As EventArgs)
        Await RunJobAsync("Evo Customer XRef to DB", AddressOf CustomerXrefJob.RunAsync)
    End Sub

    Private Async Sub RunCustomerMaster(sender As Object, e As EventArgs)
        Await RunJobAsync("QB Customer Master (Delta)", AddressOf CustomerMasterDeltaJob.RunAsync)
    End Sub

    Private Async Sub RunCustomerMasterFull(sender As Object, e As EventArgs)
        ' Confirmation added per the same "confirm before destructive actions" pattern
        ' used elsewhere in this port (e.g. TableEditorForm's delete-row confirmation) -
        ' this truncates and reloads Customer_QB entirely, which other jobs depend on
        ' (Room Usage Report, PnL to DB, Customer Master's own gallery).
        Dim confirm = MessageBox.Show(Me,
            "This will completely clear and reload the Customer_QB table from QuickBooks. Continue?",
            "QB Customer Master to DB - Full", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If confirm <> DialogResult.Yes Then Return

        Await RunJobAsync("QB Customer Master (Full)", AddressOf CustomerMasterFullJob.RunAsync)
    End Sub

    Private Async Sub RunPnLToDb(sender As Object, e As EventArgs)
        ' Defaults to day 1 and day 31 of last month - both get snapped to whole-month
        ' boundaries inside the job regardless of what's actually picked.
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(1, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(31, -1)
        Using dlg As New DateRangeDialog("PnL to DB", "From Date", "To Date", defaultFrom, defaultTo)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunJobAsync("PnL to DB", Function() PnLToDbJob.RunAsync(dlg.FromDate, dlg.ToDate))
        End Using
    End Sub

    Private Async Sub RunRoomUsageReport(sender As Object, e As EventArgs)
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using dlg As New DateRangeDialog("Room Usage Report", "Bill From Date", "Bill To Date", defaultFrom, defaultTo)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunJobAsync("Room Usage Report", Function() RoomUsageReportJob.RunAsync(dlg.FromDate, dlg.ToDate))
        End Using
    End Sub

    Private Async Sub RunSpheremailCharges(sender As Object, e As EventArgs)
        ' Same date-range UI pattern as Call Counts/Variable Charges - same
        ' 26th-of-last-month through 25th-of-this-month billing-cycle default.
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using dlg As New DateRangeDialog("Spheremail Charges to DB", "From Date", "To Date", defaultFrom, defaultTo)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunJobAsync("Spheremail Charges to DB", Function() SpheremailChargesToDbJob.RunAsync(dlg.FromDate, dlg.ToDate))
        End Using
    End Sub

    Private Async Sub RunVariableCharges(sender As Object, e As EventArgs)
        ' Same date-range UI pattern as Call Counts - same 26th-of-last-month through
        ' 25th-of-this-month billing-cycle default.
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using dlg As New DateRangeDialog("Variable Charges to DB", "From Date", "To Date", defaultFrom, defaultTo)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunJobAsync("Variable Charges to DB", Function() VariableChargesToDbJob.RunAsync(dlg.FromDate, dlg.ToDate))
        End Using
    End Sub

    Private Async Sub RunCallCounts(sender As Object, e As EventArgs)
        ' Interactive default: a fixed billing-cycle default (26th of last month through
        ' 25th of this month) - NOT the same as the batch-mode default (GetNextStartDate,
        ' which computes MAX(StartDate)+1). These are genuinely different defaults for the
        ' interactive vs. headless paths - don't unify them.
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using dlg As New DateRangeDialog("Call Counts to DB", "From Date", "To Date", defaultFrom, defaultTo)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunJobAsync("Call Counts to DB", Function() CallCountsJob.RunAsync(dlg.FromDate, dlg.ToDate))
        End Using
    End Sub

    Private Async Sub RunIncomeDb(sender As Object, e As EventArgs)
        ' Interactive default: 1st through last day of LAST month - genuinely different
        ' from the batch-mode default (current month, used by Program.vb's headless
        ' dispatch). Don't assume interactive and batch should share one default.
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(1, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(31, -1)
        Using dlg As New DateRangeDialog("Income to DB", "From Date", "To Date", defaultFrom, defaultTo)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunJobAsync("Income to DB", Function() IncomeDbJob.RunAsync(dlg.FromDate, dlg.ToDate))
        End Using
    End Sub

    Private Async Sub RunSpheremailStorage(sender As Object, e As EventArgs)
        Dim result = MessageBox.Show(Me, "Run for Individual Customers?", "Spheremail Storage Emails", MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Cancel Then Return
        Dim mode = If(result = DialogResult.Yes, SphereMailStorageEmailJob.Mode.IndividualCustomers, SphereMailStorageEmailJob.Mode.StaffSummary)
        Await RunJobAsync("Spheremail Storage Emails", Function() SphereMailStorageEmailJob.RunAsync(mode))
    End Sub

    Private Async Sub RunPaperCut(sender As Object, e As EventArgs)
        Await RunJobAsync("Papercut Scan Actions and Users", Function() Task.Run(Function() PaperCutSyncJob.Run()))
    End Sub

    Private Async Sub RunRemoteLockUsers(sender As Object, e As EventArgs)
        Await RunJobAsync("RemoteLock Users", AddressOf RemoteLockUsersJob.RunAsync)
    End Sub

    Private Async Sub RunRemoteLockAuth(sender As Object, e As EventArgs)
        Using dlg As New RemoteLockAuthDialog()
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunSelfReportingJobAsync("RemoteLock Refresh Token",
                Function() RemoteLockAuth.ExchangeAuthorizationCodeAsync(dlg.ClientId, dlg.ClientSecret, dlg.Code),
                successMessage:="completed successfully - new token saved to Config.")
        End Using
    End Sub

    Private Async Sub RunAfterHours(sender As Object, e As EventArgs)
        Await RunSelfReportingJobAsync("AfterHours Room Usage", Function() AfterHoursJob.RunAsync(1))
    End Sub

    Private Async Sub RunEarlyMeeting(sender As Object, e As EventArgs)
        ' Primarily a Task Scheduler job (see Program.vb's "EarlyMeeting" headless case) -
        ' this button exists just so the option is visible/discoverable on the dashboard,
        ' not because manual runs are the main use case. Same call signature as the
        ' headless dispatch: RunAsync() takes no arguments.
        Await RunSelfReportingJobAsync("Early Morning Meeting Emails", Function() EarlyMeetingJob.RunAsync())
    End Sub

    Private Sub RunEditAnsweringConfig(sender As Object, e As EventArgs)
        Using frm As New TableEditorForm("Answering_Config")
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunEditConfig(sender As Object, e As EventArgs)
        Using frm As New TableEditorForm("Config")
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunEditErrorLog(sender As Object, e As EventArgs)
        ' Newest-first, since this table can grow large over time - see TableEditorForm's
        ' TopRowLimit safety cap.
        Using frm As New TableEditorForm("Error_Log", orderByColumn:="[Time]")
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunEditHolidays(sender As Object, e As EventArgs)
        Using frm As New TableEditorForm("Holidays")
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunEditIoEmployees(sender As Object, e As EventArgs)
        ' Read-only - IO_Employees is actually a view, not a table.
        Using frm As New TableEditorForm("IO_Employees", isReadOnly:=True)
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunEditSendProXref(sender As Object, e As EventArgs)
        ' SendPro_XRef is a real table with primary key [Company] - editable via
        ' TableEditorForm/SqlCommandBuilder without the kind of missing-PK issue
        ' Error_Log had.
        Using frm As New TableEditorForm("SendPro_XRef")
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunEditSendPro(sender As Object, e As EventArgs)
        ' Searches/edits/deletes single SendPro rows directly against the table.
        ' Newest-first, since this table accumulates mail-forward history over time.
        ' Quick filter for Account_Num = 1 - the placeholder used when no account could
        ' be resolved automatically during import, and the main thing worth quickly
        ' finding and correcting here.
        Dim quickFilters = New List(Of (label As String, filterExpression As String)) From {
            ("Show Unresolved Accounts (Account_Num = 1)", "Account_Num = 1")
        }
        Using frm As New TableEditorForm("SendPro", orderByColumn:="Transaction_Date", quickFilters:=quickFilters)
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunCustomerMasterEditor(sender As Object, e As EventArgs)
        ' Master-detail editor for Customer_Ops_Header + Customer_Ops_Item.
        Using frm As New CustomerMasterForm()
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunRandomFacilityCode(sender As Object, e As EventArgs)
        Using frm As New RandomFacilityCodeForm()
            frm.ShowDialog(Me)
        End Using
    End Sub

End Class
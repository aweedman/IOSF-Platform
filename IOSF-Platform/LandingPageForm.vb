Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' New interactive dashboard, not a direct port of any single Access form - it exists to
''' make manual testing of the ported jobs practical. Layout mirrors the original Access
''' Landing Page's sectioned structure (Quickbooks Interfaces / SQL Server Interfaces /
''' Evo Interfaces / Reports / Other) so the two stay easy to cross-reference, including
''' every button from the original - not just the ones ported so far. Buttons without a
''' backing job yet are added disabled, with a "Not yet ported" tooltip, so this also
''' works as a visual checklist against Asana (the leading numbers on many buttons match
''' Asana step numbers, per Al - preserved here for the same reason).
'''
''' LAYOUT FIX: the outer scrollable container is a plain Panel with MANUAL vertical
''' positioning (tracked via currentY), not a FlowLayoutPanel. An earlier version nested a
''' wrapping FlowLayoutPanel (one per section) inside an OUTER FlowLayoutPanel - nesting a
''' wrap-enabled FlowLayoutPanel inside another FlowLayoutPanel's own layout engine
''' produced unpredictable width measurement (confirmed: buttons flowed off-screen
''' horizontally instead of wrapping, with no way to reach later sections). A plain outer
''' Panel with explicit Y-tracking avoids that nested-layout ambiguity entirely. Each
''' section's button grid still uses a FlowLayoutPanel (WrapContents=True) internally, but
''' with its width PINNED via matching MinimumSize/MaximumSize (both set to the same
''' width, height uncapped) - the standard, reliable WinForms way to get "fixed width,
''' auto height, wraps" instead of just growing horizontally forever.
'''
''' As more Landing Page.cls buttons get ported, find the matching placeholder button
''' below and wire it up the same way as the existing working ones - just needs a handler
''' and a call into RunJobAsync/RunSelfReportingJobAsync; no layout change needed.
'''
''' Interactive date/mode selection (DateRangeDialog, the SphereMail Yes/No prompt, the
''' RemoteLock auth dialog) lives here rather than in the job classes themselves, matching
''' the separation-of-concerns decision made throughout this port: job logic takes
''' explicit parameters, UI decides how to obtain them.
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

    ''' <summary>
    ''' Builds every section from the original Landing Page, in the same order and with
    ''' the same button labels (including Asana-referencing numbers). See class remarks.
    ''' </summary>
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
        AddButton(other, "RemoteLock Refresh Token...", AddressOf RunRemoteLockAuth)
        AddButton(other, "Spheremail Worklist", AddressOf RunSpheremailWorklist)
        AddButton(other, "Papercut Scan Actions and Users", AddressOf RunPaperCut)
        AddButton(other, "Edit Customer Master", AddressOf RunCustomerMasterEditor)
        AddButton(other, "Random Facility Code", AddressOf RunRandomFacilityCode)
        FinishSection(other)

        ' New section, not from the original Landing Page - direct table view/add/edit/
        ' delete access mirroring how Al used Access's own linked-table datasheet view for
        ' these specific tables. IO_Employees is read-only per Al (a view, not a table).
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
            If selectedType = 0 Then Return ' matches the original's own "If ttype = 0 Then Exit Sub"

            Dim defaultDate = DefaultDateHelper.ComputeDefaultDate(1, -1) ' 1st of last month, matching the original's DatePicker(FromDate, 1, -1, "From Date")
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
        Dim defaultDate = DefaultDateHelper.ComputeDefaultDate(1, 1) ' 1st of NEXT month, matching the original's own DatePicker default
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

                ' Reuses the SAME, already-proven PDF generator SphereMailStorageEmailJob
                ' uses for its email attachments - not a separate grid/print implementation.
                Dim pdfPath = IO.Path.Combine(IO.Path.GetTempPath(), $"Spheremail Storage Report {DateTime.Now:yyyyMMdd_HHmmss}.pdf")
                Await ReportGenerator.GenerateSphereMailStoragePdfAsync(locationRows, location, pdfPath)

                ' Opens directly with the system's default PDF viewer, matching the
                ' original's own DoCmd.OpenReport ..., acViewPreview - no intermediate
                ' grid step, per Al's explicit request.
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

    Private Sub RunSpheremailWorklist(sender As Object, e As EventArgs)
        Using worklistForm As New SpheremailWorklistForm()
            worklistForm.ShowDialog(Me)
        End Using
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
        ' Three dates, matching the original exactly: billing cycle start/end (same
        ' 26th-of-last-month through 25th-of-this-month defaults as other billing-cycle
        ' jobs), plus a separate Posting Date (defaults to today, matching the original's
        ' own DatePicker(InvDate, 25, 0, ...) - day 25 of THIS month, i.e. "today's
        ' billing cycle" rather than last month's).
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
            ' Per Al: logs each created Invoice/Credit Memo to the UI log as it happens.
            ' The job runs on a background thread, so this callback marshals back to the
            ' UI thread via BeginInvoke before touching logBox - AppendLog itself is not
            ' safe to call directly from a background thread.
            Dim logCallback As Action(Of String) = Sub(msg) BeginInvoke(New Action(Sub() AppendLog(msg)))
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
            Dim logCallback As Action(Of String) = Sub(msg) BeginInvoke(New Action(Sub() AppendLog(msg)))
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
        ' Matches the original's own DatePicker defaults exactly (day 1 and day 31 of
        ' last month) - both get snapped to whole-month boundaries inside the job
        ' regardless of what's actually picked, same as the original's DateSerial(...)
        ' logic.
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
        ' Same date-range UI pattern as Call Counts, per Al - same 26th-of-last-month
        ' through 25th-of-this-month billing-cycle default.
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using dlg As New DateRangeDialog("Variable Charges to DB", "From Date", "To Date", defaultFrom, defaultTo)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunJobAsync("Variable Charges to DB", Function() VariableChargesToDbJob.RunAsync(dlg.FromDate, dlg.ToDate))
        End Using
    End Sub

    Private Async Sub RunCallCounts(sender As Object, e As EventArgs)
        ' Matches the original's interactive-mode default exactly: DatePicker(BillStartDate,
        ' 26, -1, ...) / DatePicker(BillEndDate, 25, 0, ...) - a fixed billing-cycle default
        ' (26th of last month through 25th of this month), NOT the batch-mode
        ' MAX(StartDate)+1 calculation - those are two genuinely different defaults in the
        ' original for the interactive vs. headless paths. GetNextStartDate() (batch-mode)
        ' was mistakenly used here in an earlier version of this file - fixed.
        Dim defaultFrom = DefaultDateHelper.ComputeDefaultDate(26, -1)
        Dim defaultTo = DefaultDateHelper.ComputeDefaultDate(25, 0)
        Using dlg As New DateRangeDialog("Call Counts to DB", "From Date", "To Date", defaultFrom, defaultTo)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunJobAsync("Call Counts to DB", Function() CallCountsJob.RunAsync(dlg.FromDate, dlg.ToDate))
        End Using
    End Sub

    Private Async Sub RunIncomeDb(sender As Object, e As EventArgs)
        ' Matches the original's interactive-mode default exactly: DatePicker(FromDate, 1,
        ' -1, ...) / DatePicker(ToDate, 31, -1, ...) - 1st through last day of LAST month.
        ' This is genuinely different from the batch-mode default (current month, which
        ' Program.vb's headless dispatch already gets right) - same class of mistake just
        ' fixed for Call Counts: don't assume interactive and batch share one default.
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
        ' Read-only per Al - IO_Employees is actually a view, not a table.
        Using frm As New TableEditorForm("IO_Employees", isReadOnly:=True)
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunEditSendProXref(sender As Object, e As EventArgs)
        ' SendPro_XRef confirmed a real table with primary key [Company] via the
        ' Edit_SendPro PowerApp's own metadata, inspected earlier in this port - editable
        ' via TableEditorForm/SqlCommandBuilder without the kind of missing-PK issue
        ' Error_Log had.
        Using frm As New TableEditorForm("SendPro_XRef")
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunEditSendPro(sender As Object, e As EventArgs)
        ' Replaces the "Edit SendPro" PowerApp, which searched/edited/deleted single
        ' SendPro rows directly against the same table (confirmed by inspecting the
        ' .msapp package - no stored procedures involved, despite Al's initial recollection).
        ' Newest-first, since this table accumulates mail-forward history over time.
        ' Quick filter for Account_Num = 1 - the placeholder Al enters via
        ' SendProForwardsToDbJob when no account could be resolved automatically, and the
        ' main thing he needs to quickly find and correct here.
        Dim quickFilters = New List(Of (label As String, filterExpression As String)) From {
            ("Show Unresolved Accounts (Account_Num = 1)", "Account_Num = 1")
        }
        Using frm As New TableEditorForm("SendPro", orderByColumn:="Transaction_Date", quickFilters:=quickFilters)
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunCustomerMasterEditor(sender As Object, e As EventArgs)
        ' Replaces the "Customer_Master" PowerApp - a master-detail editor for
        ' Customer_Ops_Header + Customer_Ops_Item. No corresponding button existed in the
        ' original Access Landing Page's own button list, unlike Edit SendPro (which reused
        ' the existing 160.3 placeholder) - this looks like it was a standalone PowerApps
        ' tool. Placed here for now; can move if Al prefers a different location.
        Using frm As New CustomerMasterForm()
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub RunRandomFacilityCode(sender As Object, e As EventArgs)
        ' Replaces the "Random_Facility_Code" PowerApp - same situation as Customer
        ' Master, no corresponding original Access button, placed here for now.
        Using frm As New RandomFacilityCodeForm()
            frm.ShowDialog(Me)
        End Using
    End Sub

End Class
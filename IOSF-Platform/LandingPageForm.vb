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
        AddPlaceholder(sql, "180.1 - Spheremail Charges to DB")
        AddPlaceholder(sql, "160.2 FedEx Charges to DB")
        AddPlaceholder(sql, "160.3 Edit SendPro")
        AddPlaceholder(sql, "190.1 Kube Meetings to DB")
        AddButton(sql, "QB Customer Master to DB - Delta", AddressOf RunCustomerMaster)
        AddButton(sql, "Income to DB...", AddressOf RunIncomeDb)
        AddButton(sql, "190.3 - Call Counts to DB...", AddressOf RunCallCounts)
        AddPlaceholder(sql, "QB Customer Master to DB - Full")
        AddPlaceholder(sql, "PnL to DB")
        AddButton(sql, "Evo Customer XRef to DB", AddressOf RunCustomerXref)
        FinishSection(sql)

        Dim evo = AddSection("Evo Interfaces")
        AddPlaceholder(evo, "140.2 Copier to Evo")
        AddPlaceholder(evo, "150.2 Scan Extra Pages to Evo")
        AddPlaceholder(evo, "160.4 Forwards to Evo")
        AddPlaceholder(evo, "180.2 SphereMail to Evo")
        FinishSection(evo)

        Dim reports = AddSection("Reports")
        AddPlaceholder(reports, "140.1 Copier Counts")
        AddPlaceholder(reports, "180.3 - Spheremail Storage Report")
        AddPlaceholder(reports, "190.2 - Room Usage Report")
        AddPlaceholder(reports, "190.4 Call Counts")
        AddPlaceholder(reports, "Class Checks")
        AddPlaceholder(reports, "Mail Forwards")
        AddPlaceholder(reports, "IA Revenue per Customer")
        FinishSection(reports)

        Dim other = AddSection("Other")
        AddButton(other, "RemoteLock Users", AddressOf RunRemoteLockUsers)
        AddButton(other, "Spheremail Storage Emails...", AddressOf RunSpheremailStorage)
        AddButton(other, "Afterhours Room Usage Emails", AddressOf RunAfterHours)
        AddButton(other, "RemoteLock Refresh Token...", AddressOf RunRemoteLockAuth)
        AddPlaceholder(other, "Spheremail Worklist")
        AddButton(other, "Papercut Scan Actions and Users", AddressOf RunPaperCut)
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
            Await RunJobAsync("Kube Invoices to QB", Function() KubeInvoicesToQbJob.RunAsync(dlg.FileName))
        End Using
    End Sub

    Private Async Sub RunKubePaymentsToQb(sender As Object, e As EventArgs)
        Using dlg As New OpenFileDialog With {
            .Title = "Select the Excel file to process",
            .Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb",
            .Multiselect = False
        }
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await RunJobAsync("Kube Payments to QB", Function() KubePaymentsToQbJob.RunAsync(dlg.FileName))
        End Using
    End Sub

    Private Async Sub RunCustomerXref(sender As Object, e As EventArgs)
        Await RunJobAsync("Evo Customer XRef to DB", AddressOf CustomerXrefJob.RunAsync)
    End Sub

    Private Async Sub RunCustomerMaster(sender As Object, e As EventArgs)
        Await RunJobAsync("QB Customer Master (Delta)", AddressOf CustomerMasterDeltaJob.RunAsync)
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

End Class
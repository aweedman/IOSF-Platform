Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Generic, reusable table editor - the closest WinForms equivalent to Access's native
''' "open a linked table in datasheet view" behavior, which Al used directly for
''' Answering_Config, Config, Error_Log, Holidays, and (read-only) IO_Employees. A
''' DataGridView bound to a DataTable gives the same core experience: edit cells directly,
''' add a new row via the blank row at the bottom, delete via selecting a row and pressing
''' Delete - all built into the control, no custom per-table code needed.
'''
''' IMPORTANT CAVEAT: SqlCommandBuilder (used to auto-generate the INSERT/UPDATE/DELETE
''' statements on Save) requires the table to have a proper primary key - without one, the
''' grid will still load and display fine, but Save will throw at runtime. NOT verified
''' that Answering_Config/Error_Log/Holidays all have one (Config does, confirmed earlier
''' in this port - a surrogate Id column was added specifically because it originally
''' lacked one). If Save fails with a SqlCommandBuilder-related error, that's the fix
''' needed on that specific table, same pattern as Config's own PK issue earlier.
'''
''' Loads at most TopRowLimit rows, most-recently-added first where a reasonable ordering
''' column is available - a safety cap for tables like Error_Log that could otherwise grow
''' very large over time. Editing/deleting is still fully supported within the loaded set;
''' this only limits how much loads into the grid at once, not what can be modified.
''' </summary>
Public Class TableEditorForm
    Inherits Form

    Private Const TopRowLimit As Integer = 500

    Private ReadOnly tableName As String
    Private ReadOnly isReadOnly As Boolean
    Private ReadOnly orderByColumn As String

    Private grid As DataGridView
    Private searchBox As TextBox
    Private saveButton As Button
    Private refreshButton As Button
    Private deleteButton As Button
    Private statusLabel As Label
    Private adapter As SqlDataAdapter
    Private table As DataTable

    ''' <summary>
    ''' orderByColumn, if given, sorts newest-first (DESC) when applying TopRowLimit - pass
    ''' Nothing for tables with no obvious "newest" column, which just applies the row cap
    ''' without a particular ordering.
    '''
    ''' quickFilters, if given, renders one button per entry next to the search box - each
    ''' sets table.DefaultView.RowFilter directly to that entry's filter expression on
    ''' click (e.g. "Account_Num = 1"). Added because the generic substring search can't
    ''' cleanly express "exactly equals 1 in this one column" - it would also match 100,
    ''' 125, any value containing "1" in any column, etc. Pass Nothing for tables with no
    ''' such recurring, precise lookup need.
    ''' </summary>
    Public Sub New(tableName As String, Optional isReadOnly As Boolean = False, Optional orderByColumn As String = Nothing,
                   Optional quickFilters As List(Of (label As String, filterExpression As String)) = Nothing)
        Me.tableName = tableName
        Me.isReadOnly = isReadOnly
        Me.orderByColumn = orderByColumn

        Text = $"{tableName}{If(isReadOnly, " (read-only)", "")}"
        Width = 900
        Height = 600
        StartPosition = FormStartPosition.CenterScreen

        Dim searchPanel As New FlowLayoutPanel With {
            .Dock = DockStyle.Top,
            .Height = 40,
            .FlowDirection = FlowDirection.LeftToRight,
            .Padding = New Padding(8)
        }
        Dim searchLabel As New Label With {.Text = "Search:", .AutoSize = True, .Padding = New Padding(0, 6, 4, 0)}
        searchBox = New TextBox With {.Width = 300}
        AddHandler searchBox.TextChanged, AddressOf SearchTextChanged
        searchPanel.Controls.Add(searchLabel)
        searchPanel.Controls.Add(searchBox)

        If quickFilters IsNot Nothing Then
            For Each qf In quickFilters
                Dim btn As New Button With {.Text = qf.label, .AutoSize = True, .Margin = New Padding(8, 0, 0, 0)}
                Dim expression = qf.filterExpression ' capture for the closure below
                AddHandler btn.Click, Sub()
                                           searchBox.Clear() ' quick filter and free-text search would otherwise fight over RowFilter
                                           table.DefaultView.RowFilter = expression
                                       End Sub
                searchPanel.Controls.Add(btn)
            Next
        End If

        grid = New DataGridView With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = Not isReadOnly,
            .AllowUserToDeleteRows = Not isReadOnly,
            .ReadOnly = isReadOnly,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells, ' natural column widths - DataGridView shows a horizontal scrollbar automatically once total column width exceeds the visible area. Fill mode (the previous setting) actively prevents this by forcing every column to squeeze into the visible width instead.
            .SelectionMode = DataGridViewSelectionMode.CellSelect
        }
        AddHandler grid.DataError, AddressOf GridDataError

        Dim buttonPanel As New FlowLayoutPanel With {
            .Dock = DockStyle.Bottom,
            .Height = 44,
            .FlowDirection = FlowDirection.LeftToRight,
            .Padding = New Padding(8)
        }

        refreshButton = New Button With {.Text = "Refresh", .Width = 100, .Margin = New Padding(0, 0, 8, 0)}
        AddHandler refreshButton.Click, AddressOf RefreshClicked
        buttonPanel.Controls.Add(refreshButton)

        If Not isReadOnly Then
            deleteButton = New Button With {.Text = "Delete Selected Row(s)", .Width = 150, .Margin = New Padding(0, 0, 8, 0)}
            AddHandler deleteButton.Click, AddressOf DeleteSelectedClicked
            buttonPanel.Controls.Add(deleteButton)
        End If

        If Not isReadOnly Then
            saveButton = New Button With {.Text = "Save Changes", .Width = 120, .Margin = New Padding(0, 0, 8, 0)}
            AddHandler saveButton.Click, AddressOf SaveClicked
            buttonPanel.Controls.Add(saveButton)
        End If

        statusLabel = New Label With {.AutoSize = True, .Anchor = AnchorStyles.Left, .Padding = New Padding(8, 12, 0, 0)}
        buttonPanel.Controls.Add(statusLabel)

        Controls.Add(grid)
        Controls.Add(buttonPanel)
        Controls.Add(searchPanel)

        LoadData()
    End Sub

    ''' <summary>
    ''' Robust fallback alongside HideBinaryColumns - catches per-cell formatting
    ''' failures the column-level Byte() check can't (e.g. an Object-typed or mixed
    ''' column where most values are fine but one specific cell holds binary data that
    ''' DataGridView still tries to render as an image). Confirmed via a real error that
    ''' persisted (once, instead of "a handful") after the column-level fix. Suppresses
    ''' the default error dialog per the dialog's own suggestion ("To replace this default
    ''' dialog please handle the DataError event") rather than trying to pre-guess every
    ''' column/value combination that might trigger it.
    ''' </summary>
    Private Sub GridDataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        e.ThrowException = False
    End Sub

    Private Sub LoadData()
        Try
            searchBox.Clear()

            Dim sql = $"SELECT TOP {TopRowLimit} * FROM {tableName}"
            If Not String.IsNullOrEmpty(orderByColumn) Then
                sql &= $" ORDER BY {orderByColumn} DESC"
            End If

            Dim conn As New SqlConnection(ConfigHelper.ConnectionString)
            adapter = New SqlDataAdapter(sql, conn)

            If Not isReadOnly Then
                Dim builder As New SqlCommandBuilder(adapter)
                adapter.InsertCommand = builder.GetInsertCommand()
                adapter.UpdateCommand = builder.GetUpdateCommand()
                adapter.DeleteCommand = builder.GetDeleteCommand()
            End If

            table = New DataTable()
            adapter.Fill(table)
            grid.DataSource = table
            HideBinaryColumns(grid)

            statusLabel.Text = $"{table.Rows.Count} row(s) loaded" & If(table.Rows.Count = TopRowLimit, $" (capped at {TopRowLimit})", "")
        Catch ex As Exception
            MessageBox.Show(Me, $"Error loading {tableName}: {ex.Message}", "Table Editor", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' DataGridView auto-detects Byte() columns and tries to render each cell as a
    ''' picture (via ImageConverter) - if the bytes aren't a valid image format, that
    ''' throws a GDI+ "Parameter is not valid" ArgumentException per row, repeatedly.
    ''' Hiding any such column avoids the render attempt entirely - safer than guessing
    ''' which specific column is binary, since this editor opens arbitrary tables.
    ''' </summary>
    Private Shared Sub HideBinaryColumns(grid As DataGridView)
        For Each col As DataGridViewColumn In grid.Columns
            If col.ValueType Is GetType(Byte()) Then col.Visible = False
        Next
    End Sub

    Private Sub SaveClicked(sender As Object, e As EventArgs)
        Try
            grid.EndEdit()
            Dim changed = adapter.Update(table)
            statusLabel.Text = $"Saved {changed} change(s)."
        Catch ex As Exception
            MessageBox.Show(Me, $"Error saving changes to {tableName}: {ex.Message}", "Table Editor", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshClicked(sender As Object, e As EventArgs)
        LoadData()
    End Sub

    ''' <summary>
    ''' Filters across every column generically (not hardcoded to specific field names),
    ''' so this works for any table this editor opens, not just SendPro. Updates live as
    ''' the user types. Wrapped in Try/Catch since DataView.RowFilter's expression syntax
    ''' can throw on certain typed input (e.g. unbalanced brackets) - silently ignoring an
    ''' in-progress, not-yet-valid filter is better than an exception mid-keystroke.
    ''' </summary>
    Private Sub SearchTextChanged(sender As Object, e As EventArgs)
        If table Is Nothing Then Return

        Try
            Dim term = searchBox.Text
            If String.IsNullOrEmpty(term) Then
                table.DefaultView.RowFilter = String.Empty
                Return
            End If

            Dim escapedTerm = term.Replace("'", "''")
            Dim clauses = table.Columns.Cast(Of DataColumn)().
                Select(Function(c) $"Convert([{c.ColumnName}], 'System.String') LIKE '%{escapedTerm}%'")
            table.DefaultView.RowFilter = String.Join(" OR ", clauses)
        Catch
            ' Likely mid-keystroke input that isn't a valid filter expression yet - ignore
            ' and keep whatever filter was last successfully applied.
        End Try
    End Sub

    ''' <summary>
    ''' Works whether the user selected via the row header (whole-row selection) or just
    ''' clicked individual cells (SelectionMode is CellSelect, to support editing single
    ''' values naturally) - collects the distinct set of rows touched by SelectedCells
    ''' either way, rather than relying only on SelectedRows, which would miss a row
    ''' selected by clicking a cell in it rather than its header.
    ''' </summary>
    Private Sub DeleteSelectedClicked(sender As Object, e As EventArgs)
        Dim rowsToDelete = grid.SelectedCells.
            Cast(Of DataGridViewCell)().
            Select(Function(c) c.OwningRow).
            Where(Function(r) Not r.IsNewRow).
            Distinct().
            ToList()

        If rowsToDelete.Count = 0 Then
            MessageBox.Show(Me, "Select a row first (click anywhere in the row, or its row header on the left).", "Table Editor")
            Return
        End If

        Dim confirm = MessageBox.Show(Me,
            $"Delete {rowsToDelete.Count} row(s)? This removes them from the grid now - click ""Save Changes"" afterward to make it permanent, or ""Refresh"" to undo.",
            "Table Editor", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If confirm <> DialogResult.Yes Then Return

        For Each row In rowsToDelete
            grid.Rows.Remove(row)
        Next

        statusLabel.Text = $"{rowsToDelete.Count} row(s) removed from the grid - click Save Changes to make it permanent."
    End Sub

End Class
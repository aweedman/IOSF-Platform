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
    Private saveButton As Button
    Private refreshButton As Button
    Private statusLabel As Label
    Private adapter As SqlDataAdapter
    Private table As DataTable

    ''' <summary>
    ''' orderByColumn, if given, sorts newest-first (DESC) when applying TopRowLimit - pass
    ''' Nothing for tables with no obvious "newest" column, which just applies the row cap
    ''' without a particular ordering.
    ''' </summary>
    Public Sub New(tableName As String, Optional isReadOnly As Boolean = False, Optional orderByColumn As String = Nothing)
        Me.tableName = tableName
        Me.isReadOnly = isReadOnly
        Me.orderByColumn = orderByColumn

        Text = $"{tableName}{If(isReadOnly, " (read-only)", "")}"
        Width = 900
        Height = 600
        StartPosition = FormStartPosition.CenterScreen

        grid = New DataGridView With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = Not isReadOnly,
            .AllowUserToDeleteRows = Not isReadOnly,
            .ReadOnly = isReadOnly,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells, ' natural column widths - DataGridView shows a horizontal scrollbar automatically once total column width exceeds the visible area. Fill mode (the previous setting) actively prevents this by forcing every column to squeeze into the visible width instead.
            .SelectionMode = DataGridViewSelectionMode.CellSelect
        }

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
            saveButton = New Button With {.Text = "Save Changes", .Width = 120, .Margin = New Padding(0, 0, 8, 0)}
            AddHandler saveButton.Click, AddressOf SaveClicked
            buttonPanel.Controls.Add(saveButton)
        End If

        statusLabel = New Label With {.AutoSize = True, .Anchor = AnchorStyles.Left, .Padding = New Padding(8, 12, 0, 0)}
        buttonPanel.Controls.Add(statusLabel)

        Controls.Add(grid)
        Controls.Add(buttonPanel)

        LoadData()
    End Sub

    Private Sub LoadData()
        Try
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

            statusLabel.Text = $"{table.Rows.Count} row(s) loaded" & If(table.Rows.Count = TopRowLimit, $" (capped at {TopRowLimit})", "")
        Catch ex As Exception
            MessageBox.Show(Me, $"Error loading {tableName}: {ex.Message}", "Table Editor", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

End Class
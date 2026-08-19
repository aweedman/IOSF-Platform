Imports System.Windows.Forms

''' <summary>
''' Displays SpheremailWorklistJob's results and offers printing (landscape), matching
''' the original's own explicit print step - see DataTablePrinter.vb.
''' </summary>
Public Class SpheremailWorklistForm
    Inherits Form

    Private worklistTable As DataTable
    Private grid As DataGridView

    Public Sub New()
        Text = "Spheremail Worklist"
        Width = 1000
        Height = 600
        StartPosition = FormStartPosition.CenterScreen

        grid = New DataGridView With {
            .Dock = DockStyle.Fill,
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            .SelectionMode = DataGridViewSelectionMode.CellSelect
        }
        AddHandler grid.DataError, Sub(sender, e) e.ThrowException = False

        Dim buttonPanel As New FlowLayoutPanel With {.Dock = DockStyle.Bottom, .Height = 44, .Padding = New Padding(8)}
        Dim printButton As New Button With {.Text = "Print", .Width = 100}
        AddHandler printButton.Click, AddressOf PrintClicked
        buttonPanel.Controls.Add(printButton)

        ' Fill-first-then-edges rule (see other forms' remarks for why this order matters).
        Controls.Add(grid)
        Controls.Add(buttonPanel)

        LoadWorklist()
    End Sub

    Private Async Sub LoadWorklist()
        Try
            Cursor = Cursors.WaitCursor
            worklistTable = Await SpheremailWorklistJob.FetchWorklist()
            grid.DataSource = worklistTable
        Catch ex As Exception
            MessageBox.Show(Me, $"Error fetching worklist: {ex.Message}", "Spheremail Worklist", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub PrintClicked(sender As Object, e As EventArgs)
        If worklistTable Is Nothing OrElse worklistTable.Rows.Count = 0 Then
            MessageBox.Show(Me, "Nothing to print.", "Spheremail Worklist")
            Return
        End If
        Try
            DataTablePrinter.PrintWithDialog(worklistTable, "Spheremail Worklist", Me)
        Catch ex As Exception
            MessageBox.Show(Me, $"Error printing: {ex.Message}", "Spheremail Worklist", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
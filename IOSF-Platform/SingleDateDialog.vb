Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Same pattern as DateRangeDialog, but for the (so far unique) case of a job needing
''' only one date, not a range - FedEx Charges to DB is the first job in this port to
''' need this.
''' </summary>
Public Class SingleDateDialog
    Inherits Form

    Private dtpDate As DateTimePicker

    Public ReadOnly Property SelectedDate As Date
        Get
            Return dtpDate.Value.Date
        End Get
    End Property

    Public Sub New(title As String, dateLabel As String, defaultDate As Date)
        Text = title
        ClientSize = New Size(320, 130)
        StartPosition = FormStartPosition.CenterParent
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        MinimizeBox = False

        Dim lbl As New Label With {.Text = dateLabel, .Location = New Point(20, 20), .AutoSize = True}
        dtpDate = New DateTimePicker With {.Location = New Point(20, 45), .Width = 260, .Value = defaultDate}

        Dim btnOk As New Button With {.Text = "OK", .Location = New Point(115, 80), .Width = 75, .DialogResult = DialogResult.OK}
        Dim btnCancel As New Button With {.Text = "Cancel", .Location = New Point(205, 80), .Width = 75, .DialogResult = DialogResult.Cancel}

        Controls.AddRange({lbl, dtpDate, btnOk, btnCancel})
        AcceptButton = btnOk
        CancelButton = btnCancel
    End Sub

End Class
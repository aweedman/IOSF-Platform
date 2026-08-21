Imports System.Windows.Forms
Imports System.Drawing

''' <summary>Simple dialog for picking a from/to date range, used by any job that operates on a date-bounded period.</summary>
Public Class DateRangeDialog
    Inherits Form

    Private dtpFrom As DateTimePicker
    Private dtpTo As DateTimePicker

    Public ReadOnly Property FromDate As Date
        Get
            Return dtpFrom.Value.Date
        End Get
    End Property

    Public ReadOnly Property ToDate As Date
        Get
            Return dtpTo.Value.Date
        End Get
    End Property

    Public Sub New(title As String, fromLabel As String, toLabel As String, defaultFrom As Date, defaultTo As Date)
        Text = title
        ClientSize = New Size(320, 200)
        StartPosition = FormStartPosition.CenterParent
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        MinimizeBox = False

        Dim lblFrom As New Label With {.Text = fromLabel, .Location = New Point(20, 20), .AutoSize = True}
        dtpFrom = New DateTimePicker With {.Location = New Point(20, 45), .Width = 260, .Value = defaultFrom}

        Dim lblTo As New Label With {.Text = toLabel, .Location = New Point(20, 75), .AutoSize = True}
        dtpTo = New DateTimePicker With {.Location = New Point(20, 100), .Width = 260, .Value = defaultTo}

        Dim btnOk As New Button With {.Text = "OK", .Location = New Point(115, 140), .Width = 75, .DialogResult = DialogResult.OK}
        Dim btnCancel As New Button With {.Text = "Cancel", .Location = New Point(205, 140), .Width = 75, .DialogResult = DialogResult.Cancel}

        Controls.AddRange({lblFrom, dtpFrom, lblTo, dtpTo, btnOk, btnCancel})
        AcceptButton = btnOk
        CancelButton = btnCancel
    End Sub

End Class
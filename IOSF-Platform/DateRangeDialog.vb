Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Replaces the DatePicker/MsgBox-confirmation-loop pattern used interactively throughout
''' the original (e.g. Command18_Click, Command48_Click) with a plain WinForms dialog -
''' one DateTimePicker per bound, no confirmation loop needed since the picker itself
''' constrains input to valid dates.
''' </summary>
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
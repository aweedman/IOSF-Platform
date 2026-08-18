Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Direct port of frmMain.cls. The original had almost no real logic - an Exit button,
''' a Form_Load that focuses the Exit button, and a date text box that opened a custom
''' date-picker dialog (modDatePicker's InputDateField). Per the earlier review of
''' modDatePicker.bas: Access has no native date picker, so that whole module existed
''' just to work around that. WinForms has had a built-in DateTimePicker control since
''' .NET 1.0, so this needs no custom picker logic at all - just the control itself.
'''
''' This is built entirely in code (no separate .Designer.vb) so it's immediately
''' buildable without manual Visual Studio designer work. Feel free to move this to the
''' designer later if you want to visually tweak layout - functionally identical either way.
''' </summary>
Public Class MainForm
    Inherits Form

    Private WithEvents cmdExit As Button
    Private WithEvents txtDate As DateTimePicker

    Public Sub New()
        Text = "Customer Master Interface"
        ClientSize = New Size(320, 130)
        StartPosition = FormStartPosition.CenterScreen
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        MinimizeBox = False

        Dim lbl As New Label With {
            .Text = "Select a date to use on your form:",
            .Location = New Point(20, 20),
            .AutoSize = True
        }

        txtDate = New DateTimePicker With {
            .Location = New Point(20, 45),
            .Width = 260,
            .Format = DateTimePickerFormat.Short
        }

        cmdExit = New Button With {
            .Text = "Exit",
            .Location = New Point(205, 85),
            .Width = 75
        }

        Controls.Add(lbl)
        Controls.Add(txtDate)
        Controls.Add(cmdExit)

        AcceptButton = Nothing
        CancelButton = cmdExit
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmdExit.Focus() ' original: Form_Load -> cmdExit.SetFocus
    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Close() ' original: Application.Quit
    End Sub

End Class
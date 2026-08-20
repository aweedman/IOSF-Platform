Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Ports the "Option Button" form's transaction-type selection (radio buttons + OK/Cancel)
''' used by Class Checks. Matches the original's exact 9 labels/order.
''' </summary>
Public Class ClassCheckTypeDialog
    Inherits Form

    Private radios As New List(Of RadioButton)

    Public ReadOnly Property SelectedType As Integer
        Get
            For i = 0 To radios.Count - 1
                If radios(i).Checked Then Return i + 1
            Next
            Return 0
        End Get
    End Property

    Public Sub New()
        Text = "Class Checks"
        ClientSize = New Size(280, 360)
        StartPosition = FormStartPosition.CenterParent
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        MinimizeBox = False

        Dim labels = {"Bills", "Checks", "Credit Card Credits", "Credit Card Expenses",
                      "Credit Memos", "Deposits", "Invoices", "Journal Entries", "Sales Receipts"}

        Dim titleLabel As New Label With {.Text = "Type of Transaction", .Location = New Point(20, 15), .AutoSize = True}
        Controls.Add(titleLabel)

        Dim y = 45
        For Each labelText In labels
            Dim rb As New RadioButton With {.Text = labelText, .Location = New Point(30, y), .AutoSize = True}
            radios.Add(rb)
            Controls.Add(rb)
            y += 26
        Next
        radios(0).Checked = True ' matches the original's own default (Option7/"Bills" has no explicit default noted, but the first option is the natural default)

        Dim btnOk As New Button With {.Text = "OK", .Location = New Point(70, y + 15), .Width = 75, .DialogResult = DialogResult.OK}
        Dim btnCancel As New Button With {.Text = "Cancel", .Location = New Point(155, y + 15), .Width = 75, .DialogResult = DialogResult.Cancel}
        Controls.AddRange({btnOk, btnCancel})
        AcceptButton = btnOk
        CancelButton = btnCancel
    End Sub

End Class
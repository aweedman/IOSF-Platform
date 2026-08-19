Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms

''' <summary>
''' Generic, reusable helper for printing a DataTable as a simple landscape grid -
''' equal-width columns, headers repeated on every page, paginated automatically.
''' First print support in this port (needed for Spheremail Worklist, which the original
''' explicitly printed - landscape, duplex - as a physical checklist for mailroom staff).
''' Built generically so any future report needing print support can reuse this rather
''' than duplicating GDI+ pagination logic.
''' </summary>
Public Module DataTablePrinter

    Private Const RowHeight As Single = 20
    Private Const HeaderHeight As Single = 24
    Private Const FontSize As Single = 9

    ''' <summary>Shows the standard Windows print dialog, then prints table if the user confirms. landscape defaults to True, matching the original Worklist caller's own orientation.</summary>
    Public Sub PrintWithDialog(table As DataTable, documentTitle As String, owner As IWin32Window, Optional landscape As Boolean = True)
        Dim pd As New PrintDocument()
        pd.DocumentName = documentTitle
        pd.DefaultPageSettings.Landscape = landscape

        Dim rowIndex = 0
        Dim headerFont As New Font("Segoe UI", FontSize, FontStyle.Bold)
        Dim cellFont As New Font("Segoe UI", FontSize)
        Dim linePen As New Pen(Color.Black, 1)

        AddHandler pd.PrintPage, Sub(sender As Object, e As PrintPageEventArgs)
                                      Dim bounds = e.MarginBounds
                                      Dim colCount = table.Columns.Count
                                      Dim colWidth As Single = CSng(bounds.Width) / colCount

                                      ' Header row, repeated on every page
                                      Dim y As Single = bounds.Top
                                      For c = 0 To colCount - 1
                                          Dim x As Single = bounds.Left + c * colWidth
                                          e.Graphics.DrawString(table.Columns(c).ColumnName, headerFont, Brushes.Black, x + 2, y + 2)
                                      Next
                                      e.Graphics.DrawLine(linePen, bounds.Left, y + HeaderHeight, bounds.Right, y + HeaderHeight)
                                      y += HeaderHeight

                                      ' Data rows until the page is full or data runs out
                                      While rowIndex < table.Rows.Count AndAlso y + RowHeight <= bounds.Bottom
                                          Dim row = table.Rows(rowIndex)
                                          For c = 0 To colCount - 1
                                              Dim x As Single = bounds.Left + c * colWidth
                                              Dim text = If(row(c) Is DBNull.Value, "", row(c).ToString())
                                              e.Graphics.DrawString(text, cellFont, Brushes.Black, x + 2, y + 2)
                                          Next
                                          y += RowHeight
                                          rowIndex += 1
                                      End While

                                      e.HasMorePages = rowIndex < table.Rows.Count
                                  End Sub

        Using dlg As New PrintDialog With {.Document = pd, .AllowSomePages = False, .AllowSelection = False}
            If dlg.ShowDialog(owner) = DialogResult.OK Then
                pd.PrinterSettings = dlg.PrinterSettings
                pd.Print()
            End If
        End Using
    End Sub

End Module
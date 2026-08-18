Imports System.IO
Imports QuestPDF.Fluent
Imports QuestPDF.Helpers
Imports QuestPDF.Infrastructure

''' <summary>
''' Recreates the "Spheremail Storage" Access report ("Mail Storage Over 30 Days") using
''' QuestPDF, based on the report's actual layout export (reports/Spheremail Storage.bas):
'''   - Header: title, generated date/time, location, IO logo
'''   - Column headers: PMB | Customer | Date Received | Sender | Qty
'''   - Detail rows sorted by Customer, then Date Received
'''   - Footer: "Page X of Y"
'''   - Landscape Letter, ~0.25" margins, Calibri/Calibri Light fonts (matches original)
'''
''' REQUIRES the QuestPDF NuGet package, and QuestPDF.Settings.License to be set once at
''' app startup (e.g. in Sub Main) - see remarks in that entry point. QuestPDF's free
''' Community license has a revenue threshold (free under ~$1M USD annual gross revenue at
''' time of writing; check questpdf.com/license for current terms) - worth confirming this
''' applies before relying on it, that's a licensing question, not a code one.
'''
''' Header background color is an approximation of the Office theme's Accent1 (#4472C4)
''' at a light tint, read directly from the exported theme file. Adjust HeaderBackColor
''' below if it doesn't match your eye closely enough - I don't have the exact tint
''' percentage QuestPDF would need to reproduce Access's own tint math precisely.
''' </summary>
Public Module ReportGenerator

    Private Const HeaderBackColor As String = "#D9E2F3" ' approx. Accent1 (#4472C4) light tint
    Private Const MutedTextColor As String = "#808080"  ' header row label color (original: ForeColor=8355711)
    Private Const BorderColor As String = "#A6A6A6"

    Private ReadOnly LogoPath As String = Path.Combine(AppContext.BaseDirectory, "Assets", "IO_newLogo_color_hi-res_cropped.png")

    ''' <summary>
    ''' Generates the Spheremail Storage report PDF for the given rows (already filtered
    ''' by the caller - e.g. all rows for a location, or just one customer's rows) and
    ''' writes it to outputPath.
    ''' </summary>
    Public Function GenerateSphereMailStoragePdfAsync(rows As List(Of SphereMailStorageRow), location As String, outputPath As String) As Task
        Return Task.Run(Sub()
                            Dim sortedRows = rows.OrderBy(Function(r) r.Customer).ThenBy(Function(r) r.CreatedAt).ToList()
                            Dim generatedAt = DateTime.Now

                            Document.Create(Sub(container)
                                                container.Page(Sub(page)
                                                                   page.Size(PageSizes.Letter.Landscape())
                                                                   page.Margin(0.25F, Unit.Inch)
                                                                   page.PageColor(Colors.White)
                                                                   page.DefaultTextStyle(Function(x) x.FontFamily("Calibri").FontSize(11))

                                                                   page.Header().Element(Sub(header) BuildHeader(header, location, generatedAt))
                                                                   page.Content().Element(Sub(content) BuildTable(content, sortedRows))
                                                                   page.Footer().AlignCenter().Text(Sub(x)
                                                                                                        x.Span("Page ")
                                                                                                        x.CurrentPageNumber()
                                                                                                        x.Span(" of ")
                                                                                                        x.TotalPages()
                                                                                                    End Sub)
                                                               End Sub)
                                            End Sub).GeneratePdf(outputPath)
                        End Sub)
    End Function

    Private Sub BuildHeader(container As IContainer, location As String, generatedAt As DateTime)
        container.Row(Sub(row)
                          If File.Exists(LogoPath) Then
                              row.ConstantItem(50).Height(50).Image(LogoPath).FitArea()
                          End If

                          row.RelativeItem().PaddingLeft(10).Column(Sub(col)
                                                                        col.Item().Text("Mail Storage Over 30 Days").FontSize(18).FontFamily("Calibri Light")
                                                                        col.Item().Text(location).FontSize(11)
                                                                    End Sub)

                          row.ConstantItem(180).AlignRight().Column(Sub(col)
                                                                        col.Item().AlignRight().Text(generatedAt.ToString("MMMM d, yyyy"))
                                                                        col.Item().AlignRight().Text(generatedAt.ToString("h:mm:ss tt"))
                                                                    End Sub)
                      End Sub)
    End Sub

    Private Sub BuildTable(container As IContainer, rows As List(Of SphereMailStorageRow))
        container.Table(Sub(table)
                            table.ColumnsDefinition(Sub(columns)
                                                        columns.RelativeColumn(8)   ' PMB
                                                        columns.RelativeColumn(27)  ' Customer
                                                        columns.RelativeColumn(13)  ' Date Received
                                                        columns.RelativeColumn(44)  ' Sender
                                                        columns.RelativeColumn(8)   ' Qty
                                                    End Sub)

                            table.Header(Sub(header)
                                             header.Cell().Element(Function(c) HeaderCell(c)).Text("PMB").FontColor(MutedTextColor)
                                             header.Cell().Element(Function(c) HeaderCell(c)).Text("Customer").FontColor(MutedTextColor)
                                             header.Cell().Element(Function(c) HeaderCell(c)).AlignRight().Text("Date Received").FontColor(MutedTextColor)
                                             header.Cell().Element(Function(c) HeaderCell(c)).Text("Sender").FontColor(MutedTextColor)
                                             header.Cell().Element(Function(c) HeaderCell(c)).AlignRight().Text("Qty").FontColor(MutedTextColor)
                                         End Sub)

                            For Each row In rows
                                table.Cell().Element(Function(c) DetailCell(c)).Text(row.PrivateMailboxNumber)
                                table.Cell().Element(Function(c) DetailCell(c)).Text(row.Customer)
                                table.Cell().Element(Function(c) DetailCell(c)).AlignRight().Text(row.CreatedAt.ToString("M/d/yyyy"))
                                table.Cell().Element(Function(c) DetailCell(c)).Text(row.Sender)
                                table.Cell().Element(Function(c) DetailCell(c)).AlignRight().Text(row.Quantity)
                            Next
                        End Sub)
    End Sub

    Private Function HeaderCell(container As IContainer) As IContainer
        Return container.BorderBottom(1).BorderColor(BorderColor).PaddingVertical(5).PaddingHorizontal(2)
    End Function

    Private Function DetailCell(container As IContainer) As IContainer
        Return container.PaddingVertical(4).PaddingHorizontal(2)
    End Function

End Module
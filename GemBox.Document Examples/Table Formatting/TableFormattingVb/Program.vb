Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim table As New Table(document)
        table.TableFormat.AutomaticallyResizeToFitContents = False

        ' By default Table has assigned "Table Grid" style, the same as when creating it in Microsoft Word.
        ' This base style defines borders which can be removed with the following.
        table.TableFormat.Style.TableFormat.Borders.ClearBorders()

        ' Add columns with specified width.
        table.Columns.Add(New TableColumn(60))
        table.Columns.Add(New TableColumn(120))
        table.Columns.Add(New TableColumn(180))

        ' Add rows with specified height.
        table.Rows.Add(New TableRow(document) With {.RowFormat = New TableRowFormat() With {.Height = New TableRowHeight(30, TableRowHeightRule.AtLeast)}})
        table.Rows.Add(New TableRow(document) With {.RowFormat = New TableRowFormat() With {.Height = New TableRowHeight(60, TableRowHeightRule.AtLeast)}})
        table.Rows.Add(New TableRow(document) With {.RowFormat = New TableRowFormat() With {.Height = New TableRowHeight(90, TableRowHeightRule.AtLeast)}})

        For r = 0 To 2
            For c = 0 To 2

                ' Add cell.
                Dim cell As New TableCell(document)
                table.Rows(r).Cells.Add(cell)

                ' Set cell's vertical alignment.
                cell.CellFormat.VerticalAlignment = CType(r, VerticalAlignment)

                ' Add cell content.
                Dim paragraph As New Paragraph(document, $"Cell ({r + 1},{c + 1})")
                cell.Blocks.Add(paragraph)

                ' Set cell content's horizontal alignment.
                paragraph.ParagraphFormat.Alignment = CType(c, HorizontalAlignment)

                If (r + c) Mod 2 = 0 Then

                    ' Set cell's background and borders.
                    cell.CellFormat.BackgroundColor = New Color(255, 242, 204)
                    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Double, Color.Red, 1)

                End If
            Next
        Next

        document.Sections.Add(New Section(document, table))

        document.Save("Table Formatting.docx")

    End Sub
End Module
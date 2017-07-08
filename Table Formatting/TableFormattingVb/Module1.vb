Imports System
Imports GemBox.Document
Imports GemBox.Document.Tables

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        Dim table As New Table(document)
        table.TableFormat.AutomaticallyResizeToFitContents = False
        table.TableFormat.Alignment = HorizontalAlignment.Center

        table.Columns.Add(New TableColumn() With {.PreferredWidth = 50})
        table.Columns.Add(New TableColumn() With {.PreferredWidth = 80})
        table.Columns.Add(New TableColumn() With {.PreferredWidth = 110})
        table.Columns.Add(New TableColumn() With {.PreferredWidth = 140})

        Dim row As New TableRow(document)
        row.RowFormat.Height = New TableRowHeight(100, TableRowHeightRule.AtLeast)
        table.Rows.Add(row)

        Dim cell1 As New TableCell(document, New Paragraph(document, "Cell 1-1"))
        cell1.CellFormat.TextDirection = TableCellTextDirection.TopToBottom
        cell1.CellFormat.Padding = New Padding(5, 10)
        row.Cells.Add(cell1)

        Dim cell2 As New TableCell(document, New Paragraph(document, "Cell 1-2"))
        cell2.CellFormat.VerticalAlignment = VerticalAlignment.Center
        row.Cells.Add(cell2)

        row.Cells.Add(
            New TableCell(document,
                New Paragraph(document, "Cell 1-3")) With {
                    .CellFormat = New TableCellFormat() With {
                        .BackgroundColor = Color.Red
         }})

        row.Cells.Add(
            New TableCell(document,
                New Paragraph(document, "Cell 1-4") With {
                    .ParagraphFormat = New ParagraphFormat() With {
                        .Alignment = HorizontalAlignment.Center
                 }
        }) With {
            .CellFormat = New TableCellFormat() With {
                .VerticalAlignment = VerticalAlignment.Center,
                .BackgroundColor = Color.Yellow
         }})

        document.Sections.Add(New Section(document, table))

        document.Save("Table Formatting.docx")

    End Sub

End Module
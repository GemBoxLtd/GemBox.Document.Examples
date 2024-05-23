Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim rowCount As Integer = 10
        Dim columnCount As Integer = 5

        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        ' Create a table with 100% width.
        Dim table As New Table(document)
        table.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)
        section.Blocks.Add(table)

        For r As Integer = 0 To rowCount - 1

            ' Create a row and add it to table.
            Dim row As New TableRow(document)
            table.Rows.Add(row)

            For c As Integer = 0 To columnCount - 1

                ' Create a cell and add it to row.
                Dim cell As New TableCell(document)
                row.Cells.Add(cell)

                ' Create a paragraph and add it to cell.
                Dim paragraph As New Paragraph(document, $"Cell ({r + 1},{c + 1})")
                cell.Blocks.Add(paragraph)

            Next

        Next

        document.Save("Simple Table.docx")

    End Sub

End Module

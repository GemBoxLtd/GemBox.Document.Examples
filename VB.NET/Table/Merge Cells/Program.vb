Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim table As New Table(document,
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "Cell (1,1)")),
                New TableCell(document, New Paragraph(document, "Cell (1,2) -> (1,4)")) With
                {
                    .ColumnSpan = 3
                }),
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "Cell (2,1) -> (4,1)")) With
                {
                    .RowSpan = 3
                },
                New TableCell(document, New Paragraph(document, "Cell (2,2)")),
                New TableCell(document, New Paragraph(document, "Cell (2,3)")),
                New TableCell(document, New Paragraph(document, "Cell (2,4)"))),
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "Cell (3,2)")),
                New TableCell(document, New Paragraph(document, "Cell (3,3) -> (4,4)")) With
                {
                    .ColumnSpan = 2,
                    .RowSpan = 2
                }),
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "Cell (4,2)"))))

        table.TableFormat.DefaultCellPadding = New Padding(15)

        document.Sections.Add(New Section(document, table))

        document.Save("Merge Cells.docx")

    End Sub
End Module

Imports System
Imports GemBox.Document
Imports GemBox.Document.Tables

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel
        Dim table As New Table(document,
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "Cell 1-1")),
                New TableCell(document, New Paragraph(document, "Cell 1-2")),
                New TableCell(document, New Paragraph(document, "Cell 1-3")),
                New TableCell(document, New Paragraph(document, "Cell 1-4"))),
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "Cell 2-1")),
                New TableCell(document, New Paragraph(document, "Cell 2-2 -> 2-4")) With {
                     .ColumnSpan = 3
                    }),
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "Cell 3-1")),
                New TableCell(document, New Paragraph(document, "Cell 3-2 -> 5-3")) With {
                     .ColumnSpan = 2,
                     .RowSpan = 3
                    },
                New TableCell(document, New Paragraph(document, "Cell 3-4 -> 5-4")) With {
                     .RowSpan = 3
                    }),
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "Cell 4-1"))),
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "Cell 5-1"))))

        table.TableFormat.DefaultCellPadding = New Padding(10, 4)
        table.TableFormat.PreferredWidth = New TableWidth(350, TableWidthUnit.Point)

        document.Sections.Add(New Section(document, table))

        document.Save("Merge Cells.docx")

    End Sub

End Module
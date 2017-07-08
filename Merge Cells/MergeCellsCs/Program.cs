using System;
using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        Table table = new Table(document,
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "Cell 1-1")),
                new TableCell(document, new Paragraph(document, "Cell 1-2")),
                new TableCell(document, new Paragraph(document, "Cell 1-3")),
                new TableCell(document, new Paragraph(document, "Cell 1-4"))),
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "Cell 2-1")),
                new TableCell(document, new Paragraph(document, "Cell 2-2 -> 2-4"))
                {
                    ColumnSpan = 3
                }),
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "Cell 3-1")),
                new TableCell(document, new Paragraph(document, "Cell 3-2 -> 5-3"))
                {
                    ColumnSpan = 2,
                    RowSpan = 3
                },
                new TableCell(document, new Paragraph(document, "Cell 3-4 -> 5-4"))
                {
                    RowSpan = 3
                }),
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "Cell 4-1"))),
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "Cell 5-1"))));

        table.TableFormat.DefaultCellPadding = new Padding(10, 4);
        table.TableFormat.PreferredWidth = new TableWidth(350, TableWidthUnit.Point);

        document.Sections.Add(new Section(document, table));

        document.Save("Merge Cells.docx");
    }
}

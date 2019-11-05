using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var table = new Table(document,
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "Cell (1,1)")),
                new TableCell(document, new Paragraph(document, "Cell (1,2) -> (1,4)"))
                {
                    ColumnSpan = 3
                }),
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "Cell (2,1) -> (4,1)"))
                {
                    RowSpan = 3
                },
                new TableCell(document, new Paragraph(document, "Cell (2,2)")),
                new TableCell(document, new Paragraph(document, "Cell (2,3)")),
                new TableCell(document, new Paragraph(document, "Cell (2,4)"))),
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "Cell (3,2)")),
                new TableCell(document, new Paragraph(document, "Cell (3,3) -> (4,4)"))
                {
                    ColumnSpan = 2,
                    RowSpan = 2
                }),
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "Cell (4,2)"))));

        table.TableFormat.DefaultCellPadding = new Padding(15);

        document.Sections.Add(new Section(document, table));

        document.Save("Merge Cells.docx");
    }
}
using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, remove this FreeLimitReached event handler.
        ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

        int rowCount = 10;
        int columnCount = 5;

        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        // Create a table with 100% width.
        var table = new Table(document);
        table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);
        section.Blocks.Add(table);

        for (int r = 0; r < rowCount; r++)
        {
            // Create a row and add it to table.
            var row = new TableRow(document);
            table.Rows.Add(row);

            for (int c = 0; c < columnCount; c++)
            {
                // Create a cell and add it to row.
                var cell = new TableCell(document);
                row.Cells.Add(cell);

                // Create a paragraph and add it to cell.
                var paragraph = new Paragraph(document, $"Cell ({r + 1},{c + 1})");
                cell.Blocks.Add(paragraph);
            }
        }

        document.Save("Simple Table.docx");
    }
}
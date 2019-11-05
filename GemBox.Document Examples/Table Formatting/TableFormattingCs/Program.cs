using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var table = new Table(document);
        table.TableFormat.AutomaticallyResizeToFitContents = false;

        // By default Table has assigned "Table Grid" style, the same as when creating it in Microsoft Word.
        // This base style defines borders which can be removed with the following.
        table.TableFormat.Style.TableFormat.Borders.ClearBorders();

        // Add columns with specified width.
        table.Columns.Add(new TableColumn(60));
        table.Columns.Add(new TableColumn(120));
        table.Columns.Add(new TableColumn(180));

        // Add rows with specified height.
        table.Rows.Add(new TableRow(document) { RowFormat = { Height = new TableRowHeight(30, TableRowHeightRule.AtLeast) } });
        table.Rows.Add(new TableRow(document) { RowFormat = { Height = new TableRowHeight(60, TableRowHeightRule.AtLeast) } });
        table.Rows.Add(new TableRow(document) { RowFormat = { Height = new TableRowHeight(90, TableRowHeightRule.AtLeast) } });

        for (int r = 0; r < 3; r++)
            for (int c = 0; c < 3; c++)
            {
                // Add cell.
                var cell = new TableCell(document);
                table.Rows[r].Cells.Add(cell);

                // Set cell's vertical alignment.
                cell.CellFormat.VerticalAlignment = (VerticalAlignment)r;

                // Add cell content.
                var paragraph = new Paragraph(document, $"Cell ({r + 1},{c + 1})");
                cell.Blocks.Add(paragraph);

                // Set cell content's horizontal alignment.
                paragraph.ParagraphFormat.Alignment = (HorizontalAlignment)c;

                if ((r + c) % 2 == 0)
                {
                    // Set cell's background and borders.
                    cell.CellFormat.BackgroundColor = new Color(255, 242, 204);
                    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Double, Color.Red, 1);
                }
            }

        document.Sections.Add(new Section(document, table));

        document.Save("Table Formatting.docx");
    }
}
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

        Table table = new Table(document);
        table.TableFormat.AutomaticallyResizeToFitContents = false;
        table.TableFormat.Alignment = HorizontalAlignment.Center;

        table.Columns.Add(new TableColumn() { PreferredWidth = 50 });
        table.Columns.Add(new TableColumn() { PreferredWidth = 80 });
        table.Columns.Add(new TableColumn() { PreferredWidth = 110 });
        table.Columns.Add(new TableColumn() { PreferredWidth = 140 });

        TableRow row = new TableRow(document);
        row.RowFormat.Height = new TableRowHeight(100, TableRowHeightRule.AtLeast);
        table.Rows.Add(row);

        TableCell cell1 = new TableCell(document, new Paragraph(document, "Cell 1-1"));
        cell1.CellFormat.TextDirection = TableCellTextDirection.TopToBottom;
        cell1.CellFormat.Padding = new Padding(5, 10);
        row.Cells.Add(cell1);

        TableCell cell2 = new TableCell(document, new Paragraph(document, "Cell 1-2"));
        cell2.CellFormat.VerticalAlignment = VerticalAlignment.Center;
        row.Cells.Add(cell2);

        row.Cells.Add(new TableCell(document, new Paragraph(document, "Cell 1-3"))
        {
            CellFormat = new TableCellFormat()
            {
                BackgroundColor = Color.Red
            }
        });

        row.Cells.Add(new TableCell(document, new Paragraph(document, "Cell 1-4")
        {
            ParagraphFormat = new ParagraphFormat()
            {
                Alignment = HorizontalAlignment.Center
            }
        })
        {
            CellFormat = new TableCellFormat()
            {
                VerticalAlignment = VerticalAlignment.Center,
                BackgroundColor = Color.Yellow
            }
        });

        document.Sections.Add(new Section(document, table));

        document.Save("Table Formatting.docx");
    }
}

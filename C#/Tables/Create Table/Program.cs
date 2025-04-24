using GemBox.Document;
using GemBox.Document.Tables;
using System.Data;
using System.Linq;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {

        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

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

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        int rowCount = 10;
        int columnCount = 5;

        // Create DataTable with some sample data.
        var dataTable = new DataTable();
        for (int c = 0; c < columnCount; c++)
            dataTable.Columns.Add($"Column {c + 1}");
        for (int r = 0; r < rowCount; r++)
            dataTable.Rows.Add(Enumerable.Range(0, columnCount).Select(c => $"Cell ({r + 1},{c + 1})").ToArray());

        // Create new document.
        var document = new DocumentModel();

        // Create Table element from DataTable object.
        Table table = new Table(document, rowCount, columnCount,
            (int r, int c) => new TableCell(document, new Paragraph(document, dataTable.Rows[r][c].ToString())));

        // Insert first row as Table's header.
        table.Rows.Insert(0, new TableRow(document, dataTable.Columns.Cast<DataColumn>().Select(
            dataColumn => new TableCell(document, new Paragraph(document, dataColumn.ColumnName)))));

        table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);

        document.Sections.Add(new Section(document, table));

        document.Save("Insert DataTable.docx");
    }
}
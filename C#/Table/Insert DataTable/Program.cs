using GemBox.Document;
using GemBox.Document.Tables;
using System.Data;
using System.Linq;

class Program
{
    static void Main()
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

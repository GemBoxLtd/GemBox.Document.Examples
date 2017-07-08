using System;
using System.Data;
using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        int tableRowCount = 10;
        int tableColumnCount = 5;

        // Initialize DataTable
        DataTable dt = new DataTable();

        for (int i = 0; i < tableColumnCount; i++)
            dt.Columns.Add();

        for (int i = 0; i < tableRowCount; i++)
        {
            DataRow row = dt.NewRow();
            for (int j = 0; j < tableColumnCount; j++)
            {
                row[j] = string.Format("Cell {0}-{1}", i + 1, j + 1);
            }
            dt.Rows.Add(row);
        }

        // Initialize Table
        Table table = new Table(document, tableRowCount, tableColumnCount,
            (int rowIndex, int columnIndex) => new TableCell(document, new Paragraph(document, dt.Rows[rowIndex][columnIndex].ToString())));

        table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);

        document.Sections.Add(new Section(document, table));

        document.Save("Insert DataTable.docx");
    }
}

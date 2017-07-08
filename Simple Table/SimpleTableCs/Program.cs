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

        int tableRowCount = 10;
        int tableColumnCount = 5;

        Table table = new Table(document);
        table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);

        for (int i = 0; i < tableRowCount; i++)
        {
            TableRow row = new TableRow(document);
            table.Rows.Add(row);

            for (int j = 0; j < tableColumnCount; j++)
            {
                Paragraph para = new Paragraph(document, string.Format("Cell {0}-{1}", i + 1, j + 1));

                row.Cells.Add(new TableCell(document, para));
            }
        }

        document.Sections.Add(new Section(document, table));

        document.Save("Simple Table.docx");
    }
}

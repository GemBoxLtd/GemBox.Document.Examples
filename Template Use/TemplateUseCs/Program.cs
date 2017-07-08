using System;
using System.Linq;
using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("Invoice.docx");

        int numberOfItems = 10;

        // Document contains 4 tables. Each table contains some set of information.
        Table[] tables = document.GetChildElements(true, ElementType.Table).Cast<Table>().ToArray();

        // First table contains invoice number and date.
        Table invoiceTable = tables[0];
        int invoiceNumber = 14;
        // We can get and cast first paragraph using:
        // a) Linq
        invoiceTable.Rows[0].Cells[1].Blocks.Cast<Paragraph>().First().Inlines.Add(new Run(document, invoiceNumber.ToString()));
        // b) ElementCollection.Cast<TElement>(int index)
        invoiceTable.Rows[1].Cells[1].Blocks.Cast<Paragraph>(0).Inlines.Add(new Run(document, DateTime.Now.ToShortDateString()));

        // Second table contains customer data.
        Table customerTable = tables[1];
        customerTable.Rows[0].Cells[1].Blocks.Add(new Paragraph(document, "ACME Corp"));
        customerTable.Rows[1].Cells[1].Blocks.Add(new Paragraph(document, "240 Old Country Road, Springfield, IL"));
        customerTable.Rows[2].Cells[1].Blocks.Add(new Paragraph(document, "USA"));
        customerTable.Rows[3].Cells[1].Blocks.Add(new Paragraph(document, "Joe Smith"));

        // Third table contains amount and prices.
        Table mainTable = tables[2];
        // Our main table contains only one row (for one item). If we have more items then we need to clone it.
        for (int i = 1; i < numberOfItems; i++)
            mainTable.Rows.Insert(1, mainTable.Rows[1].Clone(true));

        DateTime startDate = DateTime.Now.AddDays(-numberOfItems);
        int total = 0;

        int rowIndex;
        for (rowIndex = 1; rowIndex < numberOfItems + 1; rowIndex++)
        {
            mainTable.Rows[rowIndex].Cells[0].Blocks.Add(new Paragraph(document, startDate.AddDays(rowIndex - 1).ToString("dddd, MMMM d, yyyy")));
            // We worked between 6 and 8 hours per day.
            int workHours = rowIndex % 3 + 6;
            int price = workHours * 35;
            total += price;
            mainTable.Rows[rowIndex].Cells[1].Blocks.Cast<Paragraph>(0).Inlines.Add(new Run(document, workHours.ToString()));
            mainTable.Rows[rowIndex].Cells[2].Blocks.Cast<Paragraph>(0).Inlines.Add(new Run(document, "35.00"));
            mainTable.Rows[rowIndex].Cells[3].Blocks.Cast<Paragraph>(0).Inlines.Add(new Run(document, price.ToString("0.00")));
        }

        mainTable.Rows[rowIndex].Cells[3].Blocks.Cast<Paragraph>(0).Inlines.Add(new Run(document, total.ToString("0.00")));

        // Fourth table contains notes
        Table notesTable = tables[3];
        notesTable.Rows[1].Cells[0].Blocks.Add(new Paragraph(document, "Payment via check."));

        document.Save("Template Use.docx");
    }
}

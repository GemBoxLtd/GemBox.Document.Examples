using System;
using System.Data;
using System.Linq;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("InvoiceForMailMerge.docx");

        int numberOfItems = 10;

        // Fill items (amount and prices).
        DataTable dt = new DataTable()
        {
            TableName = "Item"
        };
        dt.Columns.Add("Date", typeof(DateTime));
        dt.Columns.Add("Hours", typeof(double));
        dt.Columns.Add("Price", typeof(double));
        dt.Columns.Add("Total", typeof(double));

        DateTime startDate = DateTime.Now.AddDays(-numberOfItems);
        for (int i = 0; i < numberOfItems; i++)
        {
            // We worked between 6 and 8 hours per day.
            int workHours = i % 3 + 6;
            double totalPrice = workHours * 35;

            dt.Rows.Add(startDate.AddDays(i), workHours, 35, totalPrice);
        }

        document.MailMerge.Execute(dt);

        // Fill invoice number and date.
        int invoiceNumber = 14;
        document.MailMerge.Execute(
            new
            {
                Number = invoiceNumber,
                InvoiceDate = DateTime.Now.ToShortDateString()
            });


        // Fill customer data.
        document.MailMerge.Execute(
            new
            {
                Name = "ACME Corp",
                Address = "240 Old Country Road, Springfield, IL",
                Country = "USA",
                ContactPerson = "Joe Smith"
            });

        // Fill total.
        document.MailMerge.Execute(new { TotalPrice = dt.Rows.Cast<DataRow>().Sum(row => (double)row["Total"]) });

        // Fill notes.
        document.MailMerge.Execute(new { Notes = "Payment via check." });

        document.Save("Merge Ranges.docx");
    }
}

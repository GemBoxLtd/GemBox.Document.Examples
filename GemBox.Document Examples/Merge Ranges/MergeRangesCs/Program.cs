using System;
using System.Data;
using System.Linq;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        int numberOfItems = 10;

        var document = DocumentModel.Load("MergeRanges.docx");

        // Create DataTable as a data source for merge range.
        var table = new DataTable() { TableName = "Items" };

        table.Columns.Add("Date", typeof(DateTime));
        table.Columns.Add("Hours", typeof(int));
        table.Columns.Add("Unit", typeof(double));
        table.Columns.Add("Price", typeof(double));

        for (int rowIndex = 1; rowIndex <= numberOfItems; rowIndex++)
        {
            DateTime date = DateTime.Today.AddDays(rowIndex - numberOfItems);
            int hours = rowIndex % 3 + 6;
            double unit = 35.0;
            double price = hours * unit;

            table.Rows.Add(date, hours, unit, price);
        }

        // Execute mail merge process for "Items" merge range.
        document.MailMerge.Execute(table);

        // Execute mail merge process again for "Number", "Date" and "TotalPrice" fields.
        document.MailMerge.Execute(
            new
            {
                Number = 10203,
                Date = DateTime.Now,
                TotalPrice = table.Rows.Cast<DataRow>().Sum(row => (double)row["Price"])
            });

        document.Save("Merged Ranges Output.docx");
    }
}
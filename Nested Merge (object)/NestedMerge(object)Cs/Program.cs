using System;
using System.Linq;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("InvoiceForNestedMailMerge.docx");

        int numberOfProjects = 4;
        int itemsPerProject = 7;

        // Fill document header.
        document.MailMerge.Execute(
            new
            {
                CompanyName = "ACME Corporation",
                Address = "240 Old Country Road, Springfield, IL",
                PrintDate = DateTime.Now.ToLongDateString()
            });

        // Create data source using LINQ and anonymous types.
        var projects = Enumerable.Range(0, numberOfProjects).Select(projectIndex =>
            new
            {
                ProjectId = projectIndex + 1,
                ContactName = "John Doe",
                ProjectName = "Project " + (projectIndex + 1),
                Items = Enumerable.Range(0, itemsPerProject).Select(itemIndex =>
                    new
                    {
                        // Some random date.
                        Date = DateTime.Now.AddDays(-itemsPerProject + itemIndex),
                        Hours = ((projectIndex + 1) * itemIndex) % 3 + 6,
                        Price = 35,
                        Total = (((projectIndex + 1) * itemIndex) % 3 + 6) * 35
                    }).ToArray()
            }).ToArray();

        document.MailMerge.FieldMerging += (sender, e) =>
        {
            if (e.IsValueFound)
            {
                // Define custom formatting.
                switch (e.FieldName)
                {
                    case "Date":
                        ((Run)e.Inline).Text = ((DateTime)e.Value).ToString("dddd, MMMM d, yyyy");
                        break;
                    case "Price":
                    case "Total":
                        ((Run)e.Inline).Text = ((int)e.Value).ToString("0.00");
                        break;
                }
            }
            else if (e.RangeName == "Projects" && e.FieldName == "TotalPrice")
            {
                // For each project calculate Total.
                e.Inline = new Run(e.Document, projects[e.RecordNumber - 1].Items.Sum(item => item.Total).ToString("0.00"));
                e.Cancel = false;
            }
        };

        // Execute nested mail merge.
        document.MailMerge.Execute(projects, "Projects");

        document.Save("Nested Merge (object).docx");
    }
}

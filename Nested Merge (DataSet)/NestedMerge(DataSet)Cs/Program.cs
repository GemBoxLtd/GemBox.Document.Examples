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

        // Create data source
        DataSet ds = new DataSet();

        // Project details
        DataTable projects = new DataTable("Projects");
        projects.Columns.Add("ProjectId", typeof(int));
        projects.Columns.Add("ContactName", typeof(string));
        projects.Columns.Add("ProjectName", typeof(string));
        ds.Tables.Add(projects);

        // Item details
        DataTable items = new DataTable("Items");
        items.Columns.Add("ProjectId", typeof(int));
        items.Columns.Add("Date", typeof(DateTime));
        items.Columns.Add("Hours", typeof(double));
        items.Columns.Add("Price", typeof(double));
        items.Columns.Add("Total", typeof(double));
        ds.Tables.Add(items);

        // Add parent-child relation 
        ds.Relations.Add("Items", projects.Columns["ProjectId"], items.Columns["ProjectId"]);

        // Fill DataSource.
        for (int i = 0; i < numberOfProjects; i++)
        {
            int projectId = i + 1;
            string contactName = "John Doe";
            string projectName = "Project " + projectId;
            projects.Rows.Add(projectId, contactName, projectName);

            DateTime startDate = DateTime.Now.AddDays(-itemsPerProject);
            for (int j = 0; j < itemsPerProject; j++)
            {
                // We worked between 6 and 8 hours per day.
                int workHours = ((i + 1) * j) % 3 + 6;
                double totalPrice = workHours * 35;
                items.Rows.Add(projectId, startDate.AddDays(j), workHours, 35, totalPrice);
            }
        }

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
                        ((Run)e.Inline).Text = ((double)e.Value).ToString("0.00");
                        break;
                }
            }
            else if (e.RangeName == "Projects" && e.FieldName == "TotalPrice")
            {
                // For each project calculate Total.
                e.Inline = new Run(e.Document, ds.Tables[e.RangeName].Rows[e.RecordNumber - 1].GetChildRows("Items").Sum(item => (double)item["Total"]).ToString("0.00"));
                e.Cancel = false;
            }
        };

        // Execute nested mail merge.
        document.MailMerge.Execute(ds, null);

        document.Save("Nested Merge (DataSet).docx");
    }
}

using GemBox.Document;
using System;
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

        int numberOfProjects = 3;
        int itemsPerProject = 8;

        string projectsRangeName = "Projects";
        string itemsRangeName = "Items";

        // Create relational data.
        var projects = new DataTable(projectsRangeName);
        projects.Columns.Add("Id", typeof(int));
        projects.Columns.Add("Name", typeof(string));

        var items = new DataTable(itemsRangeName);
        items.Columns.Add("ProjectId", typeof(int));
        items.Columns.Add("Date", typeof(DateTime));
        items.Columns.Add("Hours", typeof(int));
        items.Columns.Add("Unit", typeof(double));
        items.Columns.Add("Price", typeof(double));

        // Create DataSet with parent-child relation.
        var data = new DataSet();
        data.Tables.Add(projects);
        data.Tables.Add(items);
        data.Relations.Add(itemsRangeName, projects.Columns["Id"], items.Columns["ProjectId"]);

        for (int projectIndex = 1; projectIndex <= numberOfProjects; projectIndex++)
        {
            int id = projectIndex;
            string name = $"Project {projectIndex}";

            projects.Rows.Add(id, name);

            for (int itemIndex = 1; itemIndex <= itemsPerProject; itemIndex++)
            {
                DateTime date = DateTime.Today
                    .AddMonths(projectIndex - numberOfProjects)
                    .AddDays(itemIndex - itemsPerProject);
                int hours = itemIndex % 3 + 6;
                double unit = projectIndex * 35.0;
                double price = hours * unit;

                items.Rows.Add(id, date, hours, unit, price);
            }
        }

        var document = DocumentModel.Load("MergeNestedRanges.docx");

        // Customize mail merging to achieve calculation of "TotalPrice" for each project.
        document.MailMerge.FieldMerging += (sender, e) =>
        {
            if (e.MergeContext.RangeName == "Projects" && e.FieldName == "TotalPrice")
            {
                var total = data.Tables[e.MergeContext.RangeName].Rows[e.MergeContext.RecordIndex]
                    .GetChildRows(itemsRangeName).Sum(item => (double)item["Price"]);

                var totalRun = new Run(e.Document, total.ToString("0.00"));
                totalRun.CharacterFormat = e.Field.CharacterFormat.Clone();

                e.Inline = totalRun;
                e.Cancel = false;
            }
        };

        // Execute nested mail merge.
        document.MailMerge.Execute(data, null);

        document.Save("Merged Nested Ranges Output1.docx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        int numberOfProjects = 3;
        int itemsPerProject = 8;

        // Create hierarchical data source using LINQ and anonymous types.
        var projects = Enumerable.Range(1, numberOfProjects).Select(projectIndex =>
        {
            return new
            {
                Name = $"Project {projectIndex}",
                Items = Enumerable.Range(1, itemsPerProject).Select(itemIndex =>
                {
                    DateTime date = DateTime.Today
                        .AddMonths(projectIndex - numberOfProjects)
                        .AddDays(itemIndex - itemsPerProject);
                    int hours = itemIndex % 3 + 6;
                    double unit = projectIndex * 35.0;
                    double price = hours * unit;

                    return new { Date = date, Hours = hours, Unit = unit, Price = price };
                }).ToArray()
            };
        }).ToArray();

        var document = DocumentModel.Load("MergeNestedRanges.docx");

        // Customize mail merging to achieve calculation of "TotalPrice" for each project.
        document.MailMerge.FieldMerging += (sender, e) =>
        {
            if (e.MergeContext.RangeName == "Projects" && e.FieldName == "TotalPrice")
            {
                var total = projects[e.MergeContext.RecordIndex].Items.Sum(item => item.Price);

                var totalRun = new Run(e.Document, total.ToString("0.00"));
                totalRun.CharacterFormat = e.Field.CharacterFormat.Clone();

                e.Inline = totalRun;
                e.Cancel = false;
            }
        };

        // Execute nested mail merge.
        document.MailMerge.Execute(projects, "Projects");

        document.Save("Merged Nested Ranges Output2.docx");
    }
}

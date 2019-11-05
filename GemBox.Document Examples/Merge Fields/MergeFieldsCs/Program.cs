using System;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("MergeFields.docx");

        // Create data source for mail merge process.
        var data = new
        {
            Number = 10203,
            Date = DateTime.Now,
            Company = "ACME Corp",
            Address = "240 Old Country Road, Springfield, IL",
            Country = "USA",
            FullName = "Joe Smith"
        };

        // Execute mail merge process.
        document.MailMerge.Execute(data);

        document.Save("Mail Merge Output.docx");
    }
}
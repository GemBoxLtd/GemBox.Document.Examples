using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("MergeFields.docx");

        var customer = new { CustomerName = "John", Surname = "Doe", Date = DateTime.Now };

        document.MailMerge.Execute(customer);

        document.Save("Merge Fields.docx");
    }
}

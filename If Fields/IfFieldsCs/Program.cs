using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("MergeIfFields.docx");

        var customer = new { Gender = "M", CustomerName = "John", Surname = "Doe" };

        document.MailMerge.Execute(customer);

        document.Save("If Fields.docx");
    }
}

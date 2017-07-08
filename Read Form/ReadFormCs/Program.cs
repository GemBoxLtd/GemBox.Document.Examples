using System;
using System.Globalization;
using System.Text;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("FormFilled.docx");

        // Get a snapshot of all form fields in the document.
        var formFieldsData = document.Content.FormFieldsData;

        var sb = new StringBuilder();

        // Write header.
        sb.AppendLine("Document contains following form fields:");
        sb.AppendFormat(CultureInfo.InvariantCulture,
            "{0,-16}|{1,20} = {2,-20}|({3})",
            "Type",
            '"' + "Name" + '"',
            "Value",
            "Value type").
            AppendLine().AppendLine(new string('-', 78));

        // Write type, name, value and value type of each form field in the document.
        foreach (var formFieldData in formFieldsData)
            sb.AppendFormat(CultureInfo.InvariantCulture,
                "{0,-16}|{1,20} = {2,-20}|({3})",
                formFieldData.GetType().Name,
                '"' + formFieldData.Name + '"',
                formFieldData.Value,
                formFieldData.Value != null ? formFieldData.Value.GetType().FullName : "null").
                AppendLine();

        Console.WriteLine(sb.ToString());
    }
}

using System;
using System.Text.RegularExpressions;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("CustomInvoice.pdf");
        DocumentProperties properties = document.DocumentProperties;

        // Read PDF file's properties.
        Console.WriteLine($"Author: {properties.BuiltIn[BuiltInDocumentProperty.Author]}");
        Console.WriteLine($"Created on: {properties.BuiltIn[BuiltInDocumentProperty.DateContentCreated]}");
        Console.WriteLine();

        // Read PDF file's text content and match specified regular expression.
        var text = document.Content.ToString();
        var regex = new Regex(@"(?<Hours>\d+)\s+(?<Unit>\d+\.\d{2})\s+(?<Price>\d+\.\d{2})");
        foreach (Match match in regex.Matches(text))
        {
            var groups = match.Groups;
            Console.WriteLine($"Hours={groups["Hours"]} | Unit={groups["Unit"]} | Price={groups["Price"]}");
        }
    }
}
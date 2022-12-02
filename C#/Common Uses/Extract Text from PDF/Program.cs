using System;
using System.Linq;
using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("CustomInvoice.pdf");

        // Display file's properties.
        var properties = document.DocumentProperties;
        Console.WriteLine($"Title: {properties.BuiltIn[BuiltInDocumentProperty.Title]}");
        Console.WriteLine($"Author: {properties.BuiltIn[BuiltInDocumentProperty.Author]}");
        Console.WriteLine();

        // Get paragraphs.
        var paragraphs = document.GetChildElements(true, ElementType.Paragraph).Cast<Paragraph>();

        // Get tables.
        var tables = document.GetChildElements(true, ElementType.Table).Cast<Table>();

        // Display paragraphs and tables count.
        Console.WriteLine($"Paragraph count: {paragraphs.Count()}");
        Console.WriteLine($"Table count: {tables.Count()}");
        Console.WriteLine();

        // Display first paragraph's content.
        var paragraph = paragraphs.First();
        Console.WriteLine("Paragraph content:");
        Console.WriteLine(paragraph.Content.ToString());

        // Display last table's content.
        var table = tables.Last();
        Console.WriteLine("Table content:");

        foreach (var row in table.Rows)
        {
            Console.WriteLine(new string('-', 56));
            foreach (var cell in row.Cells)
                Console.Write($"{cell.Content.ToString().TrimEnd().PadRight(13)}|");
            Console.WriteLine();
        }
    }
}
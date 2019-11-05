using System;
using System.Linq;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");

        int numberOfSections = document.Sections.Count;
        int numberOfParagraphs = document.GetChildElements(true, ElementType.Paragraph).Count();
        int numberOfRunsAndFields = document.GetChildElements(true, ElementType.Run, ElementType.Field).Count();

        var section = document.Sections[0];

        int numberOfElements = section.GetChildElements(true).Count();
        int numberOfBlocks = section.GetChildElements(true).OfType<Block>().Count();
        int numberOfInlines = section.GetChildElements(true).OfType<Inline>().Count();

        Console.WriteLine("File has:");
        Console.WriteLine($" - {numberOfSections} sections.");
        Console.WriteLine($" - {numberOfParagraphs} paragraphs.");
        Console.WriteLine($" - {numberOfRunsAndFields} runs and fields.");

        Console.WriteLine();

        Console.WriteLine("First section has:");
        Console.WriteLine($" - {numberOfElements} elements.");
        Console.WriteLine($" - {numberOfBlocks} blocks.");
        Console.WriteLine($" - {numberOfInlines} inlines.");
    }
}
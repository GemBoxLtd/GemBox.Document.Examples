using System;
using System.Linq;
using System.Text;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("Reading.docx");

        int numberOfSections = document.Sections.Count;
        int numberOfParagraphs = document.GetChildElements(true, ElementType.Paragraph).Count();
        int numberOfRunsAndFields = document.GetChildElements(true, ElementType.Run, ElementType.Field).Count();
        int numberOfInlines = document.GetChildElements(true).OfType<Inline>().Count();

        int elements = document.Sections[0].GetChildElements(true).Count();

        StringBuilder sb = new StringBuilder();
        sb.AppendLine("File has:");
        sb.AppendLine(numberOfSections + " section");
        sb.AppendLine(numberOfParagraphs + " paragraphs");
        sb.AppendLine(numberOfRunsAndFields + " runs and fields");
        sb.AppendLine(numberOfInlines + " inlines");
        sb.AppendLine("First section contains " + elements + " elements");

        Console.WriteLine(sb.ToString());
    }
}

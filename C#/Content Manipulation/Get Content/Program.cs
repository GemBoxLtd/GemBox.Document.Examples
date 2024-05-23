using GemBox.Document;
using System;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Invoice.docx");

        // Get content from each paragraph.
        foreach (Paragraph paragraph in document.GetChildElements(true, ElementType.Paragraph))
            Console.WriteLine($"Paragraph: {paragraph.Content.ToString()}");

        // Get content from each bold run.
        foreach (Run run in document.GetChildElements(true, ElementType.Run))
            if (run.CharacterFormat.Bold)
                Console.WriteLine($"Bold run: {run.Content.ToString()}");
    }
}

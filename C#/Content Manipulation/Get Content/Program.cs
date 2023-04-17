using System;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, remove this FreeLimitReached event handler.
        ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

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
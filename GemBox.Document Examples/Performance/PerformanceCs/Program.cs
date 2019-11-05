using System;
using System.Diagnostics;
using System.Linq;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // GemBox.Document has 4 working modes, each mode has the same performance and set of features.
        // Read more on: https://www.gemboxsoftware.com/document/help/html/Evaluation_and_Licensing.htm

        // Set license key to use GemBox.Document in a Free mode.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Continue to use the component in a Trial mode when free limit is reached.
        ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

        var watch = Stopwatch.StartNew();
        var document = DocumentModel.Load("lorem-ipsum-100-pages.docx", LoadOptions.DocxDefault);
        Console.WriteLine($"Load file [sec]: {watch.Elapsed.TotalSeconds}");

        watch.Restart();
        int numberOfParagraphs = document.GetChildElements(true, ElementType.Paragraph).Count();
        Console.WriteLine($"Iterate through {numberOfParagraphs} paragraphs [sec]: {watch.Elapsed.TotalSeconds}");

        watch.Restart();
        document.Save("output.docx");
        Console.WriteLine($"Save file [sec]: {watch.Elapsed.TotalSeconds}");
    }
}
using GemBox.Document;
using System;
using System.Diagnostics;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Create document.
        var document = new DocumentModel();
        var section = new Section(document);
        document.Sections.Add(section);
        for (var i = 0; i < 10000; i++)
            section.Blocks.Add(new Paragraph(document, i.ToString()));

        var stopwatch = new Stopwatch();
        stopwatch.Start();

        // Create save options.
        var saveOptions = new DocxSaveOptions();
        saveOptions.ProgressChanged += (sender, args) =>
        {
            // Cancel operation after five seconds.
            if (stopwatch.Elapsed.Seconds >= 5)
                args.CancelOperation();
        };

        try
        {
            document.Save("Cancellation.docx", saveOptions);
            Console.WriteLine("Operation fully finished");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Operation was cancelled");
        }
    }
}

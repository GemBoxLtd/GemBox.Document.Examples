using GemBox.Document;
using System;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Console.WriteLine("Creating document");

        // Create large document.
        var document = new DocumentModel();
        var section = new Section(document);
        document.Sections.Add(section);
        for (var i = 0; i < 10000; i++)
            section.Blocks.Add(new Paragraph(document, i.ToString()));

        // Create save options.
        var saveOptions = new DocxSaveOptions();
        saveOptions.ProgressChanged += (eventSender, args) =>
        {
            Console.WriteLine($"Progress changed - {args.ProgressPercentage}%");
        };

        // Save document.
        document.Save("document.docx", saveOptions);
    }
}

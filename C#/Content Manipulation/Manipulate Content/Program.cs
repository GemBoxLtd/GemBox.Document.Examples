using GemBox.Document;
using System;
using System.Linq;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
        Example3();
    }

    static void Example1()
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

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");

        // Delete 1st paragraph's inlines.
        var paragraph1 = document.Sections[0].Blocks.Cast<Paragraph>(0);
        paragraph1.Inlines.Content.Delete();

        // Delete 3rd and 4th run from the 2nd paragraph.
        var paragraph2 = document.Sections[0].Blocks.Cast<Paragraph>(1);
        var runsContent = new ContentRange(
            paragraph2.Inlines[2].Content.Start,
            paragraph2.Inlines[3].Content.End);
        runsContent.Delete();

        // Delete specified text content.
        var bracketContent = document.Content.Find("(").First();
        bracketContent.Delete();

        document.Save("Delete Content.docx");
    }

    static void Example3()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        // Create the whole document using fluent API.
        document.Content.Start
            .LoadText("First paragraph.")
            .InsertRange(new Paragraph(document, "Second paragraph.").Content)
            .LoadText("\n")
            .LoadText("Paragraph with bold text.", new CharacterFormat() { Bold = true });

        var section = document.Sections[0];

        // Prepend text to second paragraph.
        section.Blocks[1].Content.Start.LoadText(" Some Prefix ", new CharacterFormat() { Subscript = true });

        // Append text to second paragraph.
        section.Blocks[1].Content.End.LoadText(" Some Suffix ", new CharacterFormat() { Superscript = true });

        // Insert HTML paragraph before third paragraph.
        section.Blocks[2].Content.Start.LoadText("<p style='font:italic 11pt Calibri;color:royalblue;'>Paragraph from HTML content with blue and italic text.</p>",
            new HtmlLoadOptions());

        // Insert RTF paragraph after fourth paragraph.
        section.Blocks[3].Content.End.LoadText(@"{\rtf1\ansi\deff0{\colortbl ;\red255\green128\blue64;}\cf1 Paragraph from RTF content with orange text.\par\par}",
            new RtfLoadOptions());

        document.Save("Insert Content.docx");
    }
}

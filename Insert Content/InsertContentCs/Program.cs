using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        var bold = new CharacterFormat() { Bold = true };

        // Create the whole document using streamlike API.
        document.Content.Start.
            LoadText("First paragraph.").
            InsertRange(new Paragraph(document, "Second paragraph.").Content).
            LoadText("\nThird bold paragraph.", bold);

        // Prepend text to second paragraph.
        document.Sections[0].Blocks[1].Content.Start.LoadText("Some prefix (");

        // Append text to second paragraph.
        document.Sections[0].Blocks[1].Content.End.LoadText(") some suffix.");

        // Append text formatted using HTML tags.
        document.Sections[0].Blocks[2].Content.End.LoadText("<p>Fourth paragraph is added as <b>HTML</b> content.</p>", LoadOptions.HtmlDefault);

        document.Save("Insert Content.docx");
    }
}

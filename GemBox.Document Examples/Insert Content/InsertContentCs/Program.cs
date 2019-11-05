using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
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
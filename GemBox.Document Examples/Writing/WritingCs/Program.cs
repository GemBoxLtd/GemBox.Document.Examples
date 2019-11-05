using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        // Create new empty document.
        var document = new DocumentModel();

        // Add new section with two paragraphs, containing some text and symbols.
        document.Sections.Add(
            new Section(document,
                new Paragraph(document,
                    new Run(document, "This is our first paragraph with symbols added on a new line."),
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "\xFC" + "\xF0" + "\x32") { CharacterFormat = { FontName = "Wingdings", Size = 48 } }),
                new Paragraph(document, "This is our second paragraph.")));

        // Save Word document to file's path.
        document.Save("Writing1.docx");
    }

    static void Example2()
    {
        // Create new empty document.
        var document = new DocumentModel();

        // Add plain text to document.
        document.Content.LoadText("This is a plain text.", new CharacterFormat() { FontName = "Arial" });

        // Insert RTF formatted text at the beginning of the document.
        var position = document.Content.Start.LoadText(@"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial Black;}}{\colortbl ;\red255\green128\blue64;}\f0\cf1 This is rich formatted text.}",
            LoadOptions.RtfDefault);

        // Insert HTML formatted text after the previous text.
        position.LoadText("<p style='font-family:Arial Narrow;color:royalblue;'>This is another rich formatted text.</p>",
            LoadOptions.HtmlDefault);

        // Save Word document to file's path.
        document.Save("Writing2.docx");
    }
}
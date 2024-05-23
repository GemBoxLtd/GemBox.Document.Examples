using GemBox.Document;

class Program
{
    static void Main()
    {
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Create new document.
        var document = new DocumentModel();

        // Add sample text.
        document.Content.End
            .LoadText("Lorem Ipsum\n", new CharacterFormat() { FontColor = Color.Red, Bold = true })
            .LoadText("Lorem Ipsum\n", new CharacterFormat() { FontColor = Color.Green, Italic = true })
            .LoadText("Lorem Ipsum\n", new CharacterFormat() { FontColor = Color.Blue, UnderlineStyle = UnderlineType.Single });

        // Add sample image.
        document.Content.End
            .InsertRange(new Picture(document, "Dices.png").Content);

        // Save document in DOCX and PDF format.
        document.Save("output.docx");
        document.Save("output.pdf");
    }
}

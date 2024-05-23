using GemBox.Document;
using System.Linq;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        // Document's default font size is 8pt.
        document.DefaultCharacterFormat.Size = 8;

        // Style's font size is 24pt.
        var largeFont = new CharacterStyle("Large Font") { CharacterFormat = { Size = 24 } };
        document.Styles.Add(largeFont);

        var section = new Section(document);
        document.Sections.Add(section);

        var paragraph = new Paragraph(document,
            new Run(document, "Large text that has 'Large Font' style.")
            {
                CharacterFormat = { Style = largeFont }
            },
            new SpecialCharacter(document, SpecialCharacterType.LineBreak),
            new Run(document, "Medium text that has both style and direct formatting; direct formatting has precedence over style's formatting.")
            {
                CharacterFormat = { Style = largeFont, Size = 12 }
            },
            new SpecialCharacter(document, SpecialCharacterType.LineBreak),
            new Run(document, "Small text that uses document's default formatting."));

        section.Blocks.Add(paragraph);

        // Write elements resolved font size values.
        foreach (Run run in document.GetChildElements(true, ElementType.Run).ToArray())
            section.Blocks.Add(new Paragraph(document, $"Font size: {run.CharacterFormat.Size} points. Text: {run.Text}"));

        document.Save("Style Resolution.docx");
    }
}

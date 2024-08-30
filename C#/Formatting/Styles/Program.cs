using GemBox.Document;
using System.Linq;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        // Built-in styles can be created using Style.CreateStyle method.
        var titleStyle = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Title, document);

        // We can also create our own custom styles.
        var emphasisStyle = new CharacterStyle("Emphasis");
        emphasisStyle.CharacterFormat.Italic = true;

        // To use styles, we first must add them to the document.
        document.Styles.Add(titleStyle);
        document.Styles.Add(emphasisStyle);

        // Or we can use a utility method to get a built-in style or create and add a new one in a single statement.
        var strongStyle = (CharacterStyle)document.Styles.GetOrAdd(StyleTemplateType.Strong);

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "Title (Title style)") { ParagraphFormat = { Style = titleStyle } },
                new Paragraph(document,
                    new Run(document, "Text is written using Emphasis style.") { CharacterFormat = { Style = emphasisStyle } },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Text is written using Strong style.") { CharacterFormat = { Style = strongStyle } })));

        document.Save("Styles.docx");
    }

    static void Example2()
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

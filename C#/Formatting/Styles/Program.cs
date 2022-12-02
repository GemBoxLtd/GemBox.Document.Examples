using GemBox.Document;

class Program
{
    static void Main()
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
}
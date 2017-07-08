using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        // Built-in styles can be created using Style.CreateStyle() method.
        ParagraphStyle title = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Title, document);

        // We can also create our own (custom) styles.
        CharacterStyle emphasis = new CharacterStyle("Emphasis");
        emphasis.CharacterFormat.Italic = true;

        // First add style to the document, then use it.
        document.Styles.Add(title);
        document.Styles.Add(emphasis);

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "Title (Title style)")
                {
                    ParagraphFormat = new ParagraphFormat()
                    {
                        Style = title
                    }
                },
                new Paragraph(document,
                    new Run(document, "Text is written using Strong style.")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            // Or we can use utility method.
                            Style = (CharacterStyle)document.Styles.GetOrAdd(StyleTemplateType.Strong)
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Text is written using Emphasis style.")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            Style = emphasis
                        }
                    })));

        document.Save("Styles.docx");
    }
}

using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        document.Sections.Add(
            new Section(document,
                new Paragraph(document,
                    new Run(document, "English: Hello"),
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Russian: "),
                    new Run(document, new string(new char[] { '\u0417', '\u0434', '\u0440', '\u0430', '\u0432', '\u0441', '\u0442', '\u0432', '\u0443', '\u0439', '\u0442', '\u0435' })),
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Chinese: "),
                    new Run(document, new string(new char[] { '\u4f60', '\u597d' }))),
                new Paragraph(document, "In order to see Russian and Chinese characters you need to have appropriate fonts on your machine.")));

        document.Save("Writing.docx");
    }
}

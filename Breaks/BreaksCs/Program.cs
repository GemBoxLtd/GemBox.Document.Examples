using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        Section section = new Section(document,
            new Paragraph(document,
                new Run(document, "First line."),
                new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                new Run(document, "Next line."),
                new SpecialCharacter(document, SpecialCharacterType.ColumnBreak),
                new Run(document, "Next column."),
                new SpecialCharacter(document, SpecialCharacterType.PageBreak),
                new Run(document, "Next page.")));

        section.PageSetup.TextColumns = new TextColumnCollection(2);

        document.Sections.Add(section);

        document.Save("Breaks.docx");
    }
}

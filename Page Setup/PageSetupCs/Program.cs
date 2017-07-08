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
                new Run(document, "First line"),
                new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                new Run(document, "Second line"),
                new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                new Run(document, "Third line")),
            new Paragraph(document,
                new SpecialCharacter(document, SpecialCharacterType.ColumnBreak),
                new Run(document, "First line"),
                new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                new Run(document, "Second line"),
                new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                new Run(document, "Third line")));

        PageSetup pageSetup = section.PageSetup;

        // Specify text columns.
        pageSetup.TextColumns = new TextColumnCollection(2)
        {
            LineBetween = true,
            EvenlySpaced = false
        };
        pageSetup.TextColumns[0].Width = LengthUnitConverter.Convert(1, LengthUnit.Inch, LengthUnit.Point);
        pageSetup.TextColumns[1].Width = LengthUnitConverter.Convert(2.3, LengthUnit.Inch, LengthUnit.Point);

        // Specify paper type.
        pageSetup.PaperType = PaperType.A5;

        document.Sections.Add(section);

        // Specify line numbering.
        document.Sections.Add(
            new Section(document,
                new Paragraph(document,
                new Run(document, "First line"),
                new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                new Run(document, "Second line"),
                new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                new Run(document, "Third line")))
            {
                PageSetup = new PageSetup()
                {
                    PaperType = PaperType.A5,
                    LineNumberRestartSetting = LineNumberRestartSetting.NewPage
                }
            });

        document.Save("Page Setup.docx");
    }
}

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
                new Run(document, "First page"),
                new SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            new Paragraph(document,
                new Run(document, "Even page"),
                new SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            new Paragraph(document,
                new Run(document, "Odd page"),
                new SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            new Paragraph(document, "Even page"));

        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.HeaderFirst,
                new Paragraph(document,
                    new Run(document, "First Header"))));

        // Add page number.
        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.FooterFirst,
                new Paragraph(document,
                    new Run(document, "First Footer")),
                new Paragraph(document,
                    new Field(document, FieldType.Page))
                {
                    ParagraphFormat = new ParagraphFormat() { Alignment = HorizontalAlignment.Right }
                }));

        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.HeaderDefault,
                new Paragraph(document,
                    new Run(document, "Default Header"))));

        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.FooterDefault,
                new Paragraph(document,
                    new Run(document, "Default Footer")),
                 new Paragraph(document,
                    new Field(document, FieldType.Page))
                 {
                     ParagraphFormat = new ParagraphFormat() { Alignment = HorizontalAlignment.Right }
                 }));

        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.HeaderEven,
                new Paragraph(document,
                    new Run(document, "Even Header"))));

        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.FooterEven,
                new Paragraph(document,
                    new Run(document, "Even Footer")),
                new Paragraph(document,
                    new Field(document, FieldType.Page))
                {
                    ParagraphFormat = new ParagraphFormat() { Alignment = HorizontalAlignment.Right }
                }));

        document.Sections.Add(section);

        document.Save("Header and Footer.docx");
    }
}

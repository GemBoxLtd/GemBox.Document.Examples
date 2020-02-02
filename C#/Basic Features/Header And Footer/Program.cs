using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();
        document.DefaultCharacterFormat.Size = 48;

        var section = new Section(document,
            new Paragraph(document, "First page"),
            new Paragraph(document, new SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            new Paragraph(document, "Even page"),
            new Paragraph(document, new SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            new Paragraph(document, "Odd page"),
            new Paragraph(document, new SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            new Paragraph(document, "Even page"));

        document.Sections.Add(section);

        // Add default (odd) header.
        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.HeaderDefault,
                new Paragraph(document, "Default Header")));

        // Add default (odd) footer with page number.
        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.FooterDefault,
                new Paragraph(document, "Default Footer")));

        // Add first header.
        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.HeaderFirst,
                new Paragraph(document, "First Header")));

        // Add first footer.
        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.FooterFirst,
                new Paragraph(document, "First Footer"),
                new Paragraph(document,
                    new Field(document, FieldType.Page),
                    new Run(document, " of "),
                    new Field(document, FieldType.NumPages))
                {
                    ParagraphFormat = new ParagraphFormat() { Alignment = HorizontalAlignment.Right }
                }));

        // Add even header.
        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.HeaderEven,
                new Paragraph(document, "Even Header")));

        // Add even footer with page number.
        section.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.FooterEven,
                new Paragraph(document, "Even Footer"),
                new Paragraph(document,
                    new Field(document, FieldType.Page),
                    new Run(document, " of "),
                    new Field(document, FieldType.NumPages))
                {
                    ParagraphFormat = new ParagraphFormat() { Alignment = HorizontalAlignment.Right }
                }));

        document.Save("Header and Footer.docx");
    }
}
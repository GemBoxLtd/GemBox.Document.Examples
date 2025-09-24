using GemBox.Document;

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

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();
        document.DefaultCharacterFormat.Size = 28;

        // Create sections.
        var section1 = new Section(document, new Paragraph(document, "First section content")
        { ParagraphFormat = { SpaceBefore = 40, SpaceAfter = 0 } });
        var section2 = new Section(document, new Paragraph(document, "Second section content")
        { ParagraphFormat = { SpaceBefore = 40, SpaceAfter = 0 } });
        var section3 = new Section(document, new Paragraph(document, "Third section content")
        { ParagraphFormat = { SpaceBefore = 40, SpaceAfter = 0 } });

        // Add sections to the document.
        document.Sections.Add(section1);
        document.Sections.Add(section2);
        document.Sections.Add(section3);

        // Create a header in the first section.
        section1.HeadersFooters.Add(
            new HeaderFooter(document, HeaderFooterType.HeaderDefault,
                new Paragraph(document, "Shared Header (linked across sections)")));

        // Link headers in the second and third sections to the first section header.
        section2.HeadersFooters.SetLinkedToPrevious(HeaderFooterType.HeaderDefault, true);
        section3.HeadersFooters.SetLinkedToPrevious(HeaderFooterType.HeaderDefault, true);

        document.Save("Linked To Previous.docx");
    }
}

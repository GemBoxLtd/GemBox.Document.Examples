using System;
using System.Linq;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        int heading1Count = 3;
        int heading2Count = 5;

        // Create and add Heading 1 style.
        ParagraphStyle heading1Style = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading1, document);
        document.Styles.Add(heading1Style);

        // Create and add Heading 2 style.
        ParagraphStyle heading2Style = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading2, document);
        document.Styles.Add(heading2Style);

        // Create and add TOC style.
        ParagraphStyle tocHeading = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading1, document);
        tocHeading.Name = "toc";
        tocHeading.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
        document.Styles.Add(tocHeading);

        Section section = new Section(document);
        document.Sections.Add(section);

        // Add TOC title.
        section.Blocks.Add(
            new Paragraph(document, "Contents")
            {
                ParagraphFormat =
                {
                    Style = tocHeading
                }
            });

        // Create and add new TOC.
        section.Blocks.Add(
            new TableOfEntries(document, FieldType.TOC));

        section.Blocks.Add(
            new Paragraph(document,
                new SpecialCharacter(document, SpecialCharacterType.PageBreak)));

        // Add document content.
        for (int i = 0; i < heading1Count; i++)
        {
            // Heading 1
            section.Blocks.Add(
                new Paragraph(document, "Heading 1 (" + (i + 1) + ")")
                {
                    ParagraphFormat =
                    {
                        Style = heading1Style
                    }
                });

            for (int j = 0; j < heading2Count; j++)
            {
                // Heading 2
                section.Blocks.Add(
                    new Paragraph(document, String.Format("Heading 2 ({0}-{1})", i + 1, j + 1))
                    {
                        ParagraphFormat =
                        {
                            Style = heading2Style
                        }
                    });

                // Heading 2 content.
                section.Blocks.Add(
                    new Paragraph(document,
                        "GemBox.Document is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way."));
            }
        }

        // Update TOC (TOC can be updated only after all document content is added).
        var toc = (TableOfEntries)document.GetChildElements(true, ElementType.TableOfEntries).First();
        toc.Update();

        // Update TOC's page numbers.
        // NOTE: This is not necessary when printing and saving to PDF, XPS or an image format.
        // Page numbers are automatically updated in that case.
        document.GetPaginator(new PaginatorOptions() { UpdateFields = true });

        document.Save("TOC.docx");
    }
}

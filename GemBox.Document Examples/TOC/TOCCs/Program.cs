using System.Linq;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        
        int heading1Count = 3;
        int heading2Count = 5;

        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        // Create and add "Heading 1" style.
        var heading1Style = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading1, document);
        document.Styles.Add(heading1Style);

        // Create and add "Heading 2" style.
        var heading2Style = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading2, document);
        document.Styles.Add(heading2Style);

        // Create and add new TOC element.
        section.Blocks.Add(new TableOfEntries(document, FieldType.TOC));

        section.Blocks.Add(
            new Paragraph(document,
                new SpecialCharacter(document, SpecialCharacterType.PageBreak)));

        // Add document content.
        for (int i = 0; i < heading1Count; i++)
        {
            // Add "Heading 1" paragraph with Level1 for OutlineLevel. 
            section.Blocks.Add(
                new Paragraph(document, $"Heading 1 ({i + 1})") { ParagraphFormat = { Style = heading1Style } });

            for (int j = 0; j < heading2Count; j++)
            {
                // Add "Heading 2" paragraph with Level2 for OutlineLevel.
                section.Blocks.Add(
                    new Paragraph(document, $"Heading 2 ({i + 1}-{j + 1})") { ParagraphFormat = { Style = heading2Style } });

                section.Blocks.Add(
                    new Paragraph(document, "This is a paragraph.\nIt has a default BodyText for OutlineLevel.\nIt won't be listed in TOC entries."));
            }
        }

        // Get and update TOC element.
        var toc = (TableOfEntries)document.GetChildElements(true, ElementType.TableOfEntries).First();
        toc.Update();

        // Update TOC entries page numbers.
        // This is not needed when saving to PDF, XPS and image format or when printing.
        // Page numbers are automatically updated in that case.
        /* document.GetPaginator(new PaginatorOptions() { UpdateFields = true }); */

        document.Save("TOC.docx");
    }
}
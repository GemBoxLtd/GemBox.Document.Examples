using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        // First create list style.
        ListStyle numberList = new ListStyle(ListTemplateType.NumberWithDot);

        // Then set Paragraph.ListFormat.Style and ParagraphListFormat.ListLevelNumber.
        Paragraph para1 = new Paragraph(document, "First item.");
        para1.ListFormat.Style = numberList;
        para1.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = true;

        Paragraph para2 = new Paragraph(document, "Second item.");
        para2.ListFormat.Style = numberList;
        para2.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = true;

        Paragraph para2a = new Paragraph(document, "Sub item 2a.");
        para2a.ListFormat.Style = numberList;
        para2a.ListFormat.ListLevelNumber++;
        para2a.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = true;

        Paragraph para2b = new Paragraph(document, "Sub item 2b.");
        para2b.ListFormat.Style = numberList;
        para2b.ListFormat.ListLevelNumber++;
        para2b.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = true;

        Paragraph para3 = new Paragraph(document, "Third item.");
        para3.ListFormat.Style = numberList;
        para3.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = true;

        Section section = new Section(document, para1, para2, para2a, para2b, para3);
        document.Sections.Add(section);

        section.Blocks.Add(new Paragraph(document));

        // Create list style.
        ListStyle bulletList = new ListStyle(ListTemplateType.Bullet);

        // Use it.
        section.Blocks.Add(
            new Paragraph(document, "First item.")
            {
                ParagraphFormat = new ParagraphFormat() { NoSpaceBetweenParagraphsOfSameStyle = true },
                ListFormat = new ListFormat() { Style = bulletList }
            });

        section.Blocks.Add(
            new Paragraph(document, "Second item.")
            {
                ParagraphFormat = new ParagraphFormat() { NoSpaceBetweenParagraphsOfSameStyle = true },
                ListFormat = new ListFormat() { Style = bulletList }
            });

        section.Blocks.Add(
            new Paragraph(document, "Sub item 2a.")
            {
                ParagraphFormat = new ParagraphFormat() { NoSpaceBetweenParagraphsOfSameStyle = true },
                ListFormat = new ListFormat() { Style = bulletList, ListLevelNumber = 1 }
            });

        section.Blocks.Add(
            new Paragraph(document, "Item below sub item 2a.")
            {
                ParagraphFormat = new ParagraphFormat() { NoSpaceBetweenParagraphsOfSameStyle = true },
                ListFormat = new ListFormat() { Style = bulletList, ListLevelNumber = 2 }
            });

        section.Blocks.Add(
            new Paragraph(document, "Third item.")
            {
                ParagraphFormat = new ParagraphFormat() { NoSpaceBetweenParagraphsOfSameStyle = true },
                ListFormat = new ListFormat() { Style = bulletList }
            });

        document.Save("Lists.docx");
    }
}

using GemBox.Document;
using System.Linq;

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

        var section = new Section(document);
        document.Sections.Add(section);

        var blocks = section.Blocks;

        // Create bullet list style.
        ListStyle bulletList = new ListStyle(ListTemplateType.Bullet);
        bulletList.ListLevelFormats[0].ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = true;
        bulletList.ListLevelFormats[0].CharacterFormat.FontColor = Color.Red;

        // Create bullet list items.
        blocks.Add(new Paragraph(document, "First item.")
        {
            ListFormat = { Style = bulletList }
        });
        blocks.Add(new Paragraph(document, "Second item.")
        {
            ListFormat = { Style = bulletList }
        });
        blocks.Add(new Paragraph(document, "Third item.")
        {
            ListFormat = { Style = bulletList }
        });

        blocks.Add(new Paragraph(document));

        // Create number list style.
        var numberList = new ListStyle(ListTemplateType.NumberWithDot);

        // Create number list items.
        blocks.Add(new Paragraph(document, "First item.")
        {
            ListFormat = { Style = numberList }
        });
        blocks.Add(new Paragraph(document, "Sub item 1. a.")
        {
            ListFormat = { Style = numberList, ListLevelNumber = 1 }
        });
        blocks.Add(new Paragraph(document, "Item below sub item 1. a.")
        {
            ListFormat = { Style = numberList, ListLevelNumber = 2 }
        });
        blocks.Add(new Paragraph(document, "Sub item 1. b.")
        {
            ListFormat = { Style = numberList, ListLevelNumber = 1 }
        });
        blocks.Add(new Paragraph(document, "Second item.")
        {
            ListFormat = { Style = numberList }
        });

        document.Save("Lists.docx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        int listItemsCount = 12;

        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        // Create number list style.
        var numberList = new ListStyle(ListTemplateType.NumberWithDot);

        // Customize list level formats.
        for (int level = 0; level < numberList.ListLevelFormats.Count; level++)
        {
            ListLevelFormat levelFormat = numberList.ListLevelFormats[level];

            levelFormat.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = true;
            levelFormat.Alignment = HorizontalAlignment.Left;
            levelFormat.NumberStyle = NumberStyle.Decimal;

            levelFormat.NumberPosition = 18 * level;
            levelFormat.NumberFormat = string.Concat(Enumerable.Range(1, level + 1).Select(i => $"%{i}."));
        }

        // Create number list items.
        for (int i = 0; i < listItemsCount; i++)
        {
            var paragraph = new Paragraph(document, "Lorem ipsum");

            paragraph.ListFormat.Style = numberList;
            paragraph.ListFormat.ListLevelNumber = i % 9;

            section.Blocks.Add(paragraph);
        }

        document.Save("CustomizedList.docx");
    }
}

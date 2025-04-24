using GemBox.Document;
using GemBox.Document.Tables;

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

        var bookmarkName = "TopOfDocument";

        document.Sections.Add(
            new Section(document,
                new Paragraph(document,
                    new BookmarkStart(document, bookmarkName),
                    new Run(document, "This is a 'TopOfDocument' bookmark."),
                    new BookmarkEnd(document, bookmarkName)),
                new Paragraph(document,
                    new Run(document, "The following is a link to "),
                    new Hyperlink(document, "https://www.gemboxsoftware.com/document", "GemBox.Document Overview"),
                    new Run(document, " page.")),
                 new Paragraph(document,
                    new SpecialCharacter(document, SpecialCharacterType.PageBreak),
                    new Run(document, "This is a document's second page."),
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Hyperlink(document, bookmarkName, "Return to 'TopOfDocument'.") { IsBookmarkLink = true })));

        document.Save("Bookmarks and Hyperlinks.docx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("BookmarksTemplate.docx");

        // Replace bookmark's content with plain text.
        document.Bookmarks["Company"].GetContent(false).LoadText("Acme Corporation");

        // Replace bookmark's content with HTML text.
        document.Bookmarks["Address"].GetContent(false).LoadText(
            "<span style='font: italic 8pt Calibri; color: red;'>240 Old Country Road, Springfield, IL</span>",
            LoadOptions.HtmlDefault);

        // Insert hyperlink into a bookmark.
        var hyperlink = new Hyperlink(document, "mailto:joe.doe@acme.co", "joe.doe@acme.co");
        document.Bookmarks["Contact"].GetContent(false).Set(hyperlink.Content);

        // Insert image into a bookmark.
        var picture = new Picture(document, "Acme.png");
        document.Bookmarks["Logo"].GetContent(false).Set(picture.Content);

        // Insert text and table into a bookmark.
        ContentRange itemsRange = document.Bookmarks["Items"].GetContent(false);
        itemsRange = itemsRange.LoadText("Storage:");
        var table = new Table(document, 6, 3, (r, c) => new TableCell(document, new Paragraph(document, $"Item {r}-{c}")));
        itemsRange.End.InsertRange(table.Content);

        document.Save("Modified Bookmarks.docx");
    }
}

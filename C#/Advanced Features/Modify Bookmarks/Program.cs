using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main()
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
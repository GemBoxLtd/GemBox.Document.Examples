using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
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
}
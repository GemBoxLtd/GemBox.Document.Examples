using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        string bookmarkName = "LinkToTop";

        document.Sections.Add(
            new Section(document,
                new Paragraph(document,
                    new BookmarkStart(document, bookmarkName),
                    new Run(document, "GemBox.Document"),
                    new BookmarkEnd(document, bookmarkName)),
                new Paragraph(document,
                    new Hyperlink(document, "http://www.gemboxsoftware.com/document/overview", "GemBox.Document"),
                    new Run(document, " is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way.")),
                 new Paragraph(document),
                 new Paragraph(document),
                 new Paragraph(document),
                 new Paragraph(document,
                     // When user clicks on this link it will jump to the text between BookmarkStart and BookmarkEnd.
                     new Hyperlink(document, bookmarkName, "To Top")
                     {
                         IsBookmarkLink = true
                     })));

        document.Save("Bookmarks and Hyperlinks.docx");
    }
}

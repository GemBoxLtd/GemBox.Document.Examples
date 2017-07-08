using System;
using System.IO;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("BookmarksTemplate.docx");

        string pathToResources = "Resources";

        document.Bookmarks["CompanyName"].GetContent(false).LoadText("ACME Corp");
        document.Bookmarks["CompanyAddress"].GetContent(false).LoadText("240 Old Country Road, Springfield, IL");
        document.Bookmarks["Country"].GetContent(false).LoadText("USA");
        document.Bookmarks["ContactPerson"].GetContent(false).LoadText("Joe Smith");

        document.Bookmarks["Text"].GetContent(false).LoadText(
            "GemBox.Document is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way.",
            new CharacterFormat() { Size = 14 });

        Picture picture = new Picture(document, Path.Combine(pathToResources, "Acme.png"));
        document.Bookmarks["Logo"].GetContent(false).Set(picture.Content);

        document.Save("Modify Bookmarks.docx");
    }
}

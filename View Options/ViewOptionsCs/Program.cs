using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "GemBox.Document is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way.")));

        document.ViewOptions.ViewType = ViewType.Print;
        document.ViewOptions.Zoom = 75;

        document.Save("View Options.docx");
    }
}

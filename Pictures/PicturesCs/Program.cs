using System;
using System.IO;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        string pathToResources = "Resources";

        var section = new Section(document);
        document.Sections.Add(section);

        Paragraph paragraph = new Paragraph(document);
        section.Blocks.Add(paragraph);

        Picture picture1 = new Picture(document, Path.Combine(pathToResources, "Zahnrad.gif"), 61, 53, LengthUnit.Pixel);
        paragraph.Inlines.Add(picture1);

        paragraph.Inlines.Add(new Run(document, "GemBox.Document is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way."));

        Picture picture2 = new Picture(document, Path.Combine(pathToResources, "Dices.png"));
        picture2.Layout = new FloatingLayout(
            new HorizontalPosition(HorizontalPositionType.Left, HorizontalPositionAnchor.Page),
            new VerticalPosition(2, LengthUnit.Inch, VerticalPositionAnchor.Page),
            picture2.Layout.Size) { WrappingStyle = TextWrappingStyle.InFrontOfText };
        paragraph.Inlines.Add(picture2);

        Picture picture3 = new Picture(document, Path.Combine(pathToResources, "Graphics1.wmf"), 378, 189, LengthUnit.Pixel);
        picture3.Layout = new FloatingLayout(
            new HorizontalPosition(3.5, LengthUnit.Inch, HorizontalPositionAnchor.Page),
            new VerticalPosition(2, LengthUnit.Inch, VerticalPositionAnchor.Page),
            picture3.Layout.Size) { WrappingStyle = TextWrappingStyle.BehindText };
        paragraph.Inlines.Add(picture3);

        document.Save("Pictures.docx");
    }
}

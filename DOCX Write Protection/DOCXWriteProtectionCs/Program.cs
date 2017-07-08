using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        Paragraph paragraph1 = new Paragraph(document, "This document has been opened in read-only mode. Changes cannot be made to the original document. To save changes, create a new copy of the document.");
        section.Blocks.Add(paragraph1);

        Paragraph paragraph2 = new Paragraph(document, "To enable modifying use password: 1234");
        section.Blocks.Add(paragraph2);

        WriteProtection protection = document.WriteProtection;
        // For DOCX file format: disallow resaving the document using the same file name.
        protection.SetPassword("1234");

        document.Save("DOCX Write Protection.docx");
    }
}

using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "This document has been opened in read-only mode."),
                new Paragraph(document, "Changes cannot be made to the original document."),
                new Paragraph(document, "To save changes a new copy of the document must be created.")));

        WriteProtection protection = document.WriteProtection;
        protection.SetPassword("pass");

        document.Save("DOCX Write Protection.docx");
    }
}
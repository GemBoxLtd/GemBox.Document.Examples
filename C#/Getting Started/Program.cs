using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        Section section = new Section(document);
        document.Sections.Add(section);

        Paragraph paragraph = new Paragraph(document);
        section.Blocks.Add(paragraph);

        Run run = new Run(document, "Hello World!");
        paragraph.Inlines.Add(run);

        document.Save("HelloWorld.docx");
    }
}
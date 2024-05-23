using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        // Add '{ AUTHOR }' field.
        section.Blocks.Add(
            new Paragraph(document,
                new Run(document, "Author: "),
                new Field(document, FieldType.Author, null, "Mario at GemBox")));

        // Add '{ DATE }' field.
        section.Blocks.Add(
            new Paragraph(document,
                new Run(document, "Date: "),
                new Field(document, FieldType.Date)));

        // Add '{ DATE \@ "dddd, MMMM dd, yyyy" }' field.
        section.Blocks.Add(
            new Paragraph(document,
                new Run(document, "Date with specified format: "),
                new Field(document, FieldType.Date, "\\@ \"dddd, MMMM dd, yyyy\"")));

        document.Save("Fields.docx");
    }
}

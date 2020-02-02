using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var destinationDocument = DocumentModel.Load("Invoice.docx");
        var sourceDocument = DocumentModel.Load("Reading.docx");

        // 1. The following is the easiest way how to import content from source to destination,
        // it will automatically imports the document's elements.
        //destinationDocument.Content.End.InsertRange(sourceDocument.Content);

        // 2. The following enables various customization.
        // For instance, we can change that the imported content doesn't start from new page
        // by setting "SectionStart.Continuous" on first destination section.

        // Reuse the same mapping for importing to improve performance.
        var mapping = new ImportMapping(sourceDocument, destinationDocument, false);

        // Import all sections from source document to destination document.
        foreach (var sourceSection in sourceDocument.Sections)
        {
            var destinationSection = destinationDocument.Import(sourceSection, true, mapping);
            destinationDocument.Sections.Add(destinationSection);
        }

        destinationDocument.Save("Importing.docx");
    }
}
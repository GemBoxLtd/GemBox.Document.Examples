using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load Word document, preservation feature is enabled by default.
        var document = DocumentModel.Load("Preservation.docx");

        // Modify Word document.
        document.Sections[0].Blocks.Insert(0,
            new Paragraph(document, "You can preserve unsupported features when modifying a document!"));

        // Save Word document to output file of same format together with
        // preserved information (unsupported features) from input file.
        document.Save("PreservedOutput.docx");
    }
}
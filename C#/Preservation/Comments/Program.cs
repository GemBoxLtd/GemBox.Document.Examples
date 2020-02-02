using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load Word document with preservation feature enabled.
        var loadOptions = new DocxLoadOptions() { PreserveUnsupportedFeatures = true };
        var document = DocumentModel.Load("Comments.docx", loadOptions);

        // Save Word document to output file of same format together with
        // preserved information (unsupported features) from input file.
        document.Save("Comments Output.docx");
    }
}
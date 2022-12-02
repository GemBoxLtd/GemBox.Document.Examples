using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load Word document with preservation feature enabled.
        var loadOptions = new DocxLoadOptions() { PreserveUnsupportedFeatures = true };
        var document = DocumentModel.Load("Macros.docm", loadOptions);

        // Save Word document to output file of same format together with
        // preserved information (unsupported features) from input file.
        document.Save("Preserved Output.docm");
    }
}
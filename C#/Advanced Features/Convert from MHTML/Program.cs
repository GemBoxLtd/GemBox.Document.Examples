using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load input MHTML file.
        DocumentModel document = DocumentModel.Load("Input.mhtml");

        // Save output PDF file.
        document.Save("Output.pdf");
    }
}

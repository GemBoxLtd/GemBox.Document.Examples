using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, remove this FreeLimitReached event handler.
        ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

        // Load input MHTML file.
        DocumentModel document = DocumentModel.Load("Input.mhtml");

        // Save output PDF file.
        document.Save("Output.pdf");
    }
}

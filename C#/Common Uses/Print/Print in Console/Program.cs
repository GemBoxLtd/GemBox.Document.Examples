using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("Print.docx");

        // Set Word document's page options.
        foreach (Section section in document.Sections)
        {
            PageSetup pageSetup = section.PageSetup;
            pageSetup.Orientation = Orientation.Landscape;
            pageSetup.LineNumberRestartSetting = LineNumberRestartSetting.NewPage;
            pageSetup.LineNumberDistanceFromText = 50;

            PageMargins pageMargins = pageSetup.PageMargins;
            pageMargins.Top = 20;
            pageMargins.Left = 100;
        }

        // Print Word document to default printer.
        string printer = null;
        document.Print(printer);
    }
}
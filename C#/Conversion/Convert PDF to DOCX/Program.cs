using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("CustomInvoice.pdf",
            new PdfLoadOptions()
            {
                LoadType = PdfLoadType.HighFidelity
            });

        document.Save("ConvertedFromPdf.docx");
    }
}

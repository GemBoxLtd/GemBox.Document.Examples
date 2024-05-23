using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("MergeBarcodes.docx");

        // Create data source for mail merge process.
        var data = new
        {
            QrCode = "QR Code created with GemBox.Document",
            Code128 = "1234567890",
            Ean13 = "5901234123457"
        };

        // Execute mail merge process.
        document.MailMerge.Execute(data);

        document.Save("Barcodes Merge Output.docx");
    }
}

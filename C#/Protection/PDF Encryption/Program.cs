using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        string userPassword = "pass";
        string ownerPassword = "";
        PdfPermissions permissions = PdfPermissions.None;

        var document = DocumentModel.Load("Reading.docx");

        var options = new PdfSaveOptions()
        {
            DocumentOpenPassword = userPassword,
            PermissionsPassword = ownerPassword,
            Permissions = permissions
        };

        document.Save("PDF Encryption.pdf", options);
    }
}

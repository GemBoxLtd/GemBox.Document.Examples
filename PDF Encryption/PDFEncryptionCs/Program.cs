using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        string inputPassword = "inpass";
        string outputPassword = "outpass";
        string ownerPassword = "";

        var document = DocumentModel.Load("PdfEncryption.pdf", new PdfLoadOptions() { Password = inputPassword });
        var options = new PdfSaveOptions()
        {
            DocumentOpenPassword = outputPassword,
            PermissionsPassword = ownerPassword,
            Permissions = PdfPermissions.None
        };
    }
}

using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        string inputPassword = "inpass";
        string outputPassword = "outpass";

        var document = DocumentModel.Load("DocxEncryption.docx",
            new DocxLoadOptions() { Password = inputPassword });

        document.Save("DOCX Encryption Output.docx",
            new DocxSaveOptions() { Password = outputPassword });
    }
}
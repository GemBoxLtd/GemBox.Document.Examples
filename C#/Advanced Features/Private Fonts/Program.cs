using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        
        FontSettings.FontsBaseDirectory = ".";

        var document = new DocumentModel();

        document.DefaultCharacterFormat.FontName = "Almonte Snow";
        document.DefaultCharacterFormat.Size = 48;

        document.Content.LoadText("Hello World!");

        document.Save("Private Fonts.pdf");
    }
}
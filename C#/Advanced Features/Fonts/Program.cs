using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        
        // Set the directory path where the component will look for additional font files.
        // The "." targets the current directory, so besides the installed fonts,
        // the component will be able to use the fonts within the specified directory.
        FontSettings.FontsBaseDirectory = ".";

        var document = new DocumentModel();

        document.DefaultCharacterFormat.FontName = "Almonte Snow";
        document.DefaultCharacterFormat.Size = 48;

        document.Content.LoadText("Hello World!");

        document.Save("Private Fonts.pdf");
    }
}

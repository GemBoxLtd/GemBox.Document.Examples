using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        string pathToResources = "Resources";

        FontSettings.FontsBaseDirectory = pathToResources;

        document.DefaultCharacterFormat = new CharacterFormat()
        {
            FontName = "Almonte Snow",
            Size = 16
        };

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "Hello World!")));

        document.Save("Private Fonts.pdf");
    }
}

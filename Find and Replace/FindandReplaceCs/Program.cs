using System;
using System.Linq;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("Reading.docx");

        // Find and count text.
        int documentCount = document.Content.Find("GemBox.Document").Count();

        int counter = documentCount;

        // Find text and load another text in its place.
        foreach (ContentRange item in document.Content.Find("GemBox.Document").Reverse())
            item.LoadText(string.Format("GBD ({0}/{1})", counter--, documentCount));

        // Find and replace text.
        document.Content.Replace(".NET", "C# / VB.NET");

        document.Save("Find and Replace.docx");
    }
}

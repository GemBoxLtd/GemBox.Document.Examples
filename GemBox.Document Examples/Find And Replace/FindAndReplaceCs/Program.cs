using System.Linq;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");

        // 1. The easiest way how to find and replace text.
        document.Content.Replace(".NET", "C# / VB.NET");

        // 2. You can also find and highlight text by specifying the format of replacement text.
        var highlightText = "read, write, convert and print";
        document.Content.Replace(highlightText, highlightText);
        document.Content.Replace(highlightText, highlightText,
            new CharacterFormat() { HighlightColor = Color.Yellow });

        // 3. You can also search for specified text and achieve more complex replacements.
        // Notice the "Reverse" method usage for avoiding any possible invalid state due to replacements inside iteration.
        var searchText = "GemBox.Document";
        foreach (ContentRange searchedContent in document.Content.Find(searchText).Reverse())
        {
            var replaceText = "Word library from GemBox called ";
            ContentRange replacedContent = searchedContent.LoadText(replaceText,
                new CharacterFormat() { FontColor = new Color(237, 125, 49) });

            var hyperlink = new Hyperlink(document, "https://www.gemboxsoftware.com/document", "GemBox.Document");
            replacedContent.End.InsertRange(hyperlink.Content);
        }

        document.Save("Find and Replace.docx");
    }
}
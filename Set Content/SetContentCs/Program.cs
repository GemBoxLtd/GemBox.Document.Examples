using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        // Set the content for the whole document
        document.Content.LoadText("Paragraph 1\nParagraph 2\nParagraph 3\nParagraph 4\nParagraph 5");

        var bold = new CharacterFormat()
        {
            Bold = true
        };

        // Set the content for the 2nd paragraph
        document.Sections[0].Blocks[1].Content.LoadText("Bold paragraph 2", bold);

        // Set the content for 3rd and 4th paragraph to be the same as the content of 1st and 2nd paragraph
        var para3 = document.Sections[0].Blocks[2];
        var para4 = document.Sections[0].Blocks[3];
        var destinationRange = new ContentRange(para3.Content.Start, para4.Content.End);
        var para1 = document.Sections[0].Blocks[0];
        var para2 = document.Sections[0].Blocks[1];
        var sourceRange = new ContentRange(para1.Content.Start, para2.Content.End);
        destinationRange.Set(sourceRange);

        // Set content using HTML tags
        document.Sections[0].Blocks[4].Content.LoadText("Paragraph 5 <b>(part of this paragraph is bold)</b>", LoadOptions.HtmlDefault);

        document.Save("Set Content.docx");
    }
}

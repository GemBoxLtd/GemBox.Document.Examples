using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("Reading.docx");

        // Delete paragraph break between 1st and 2nd paragraph (concatenate 1st and 2nd paragraph)
        var blocks = document.Sections[0].Blocks;
        var paragraphBreakRange = new ContentRange(blocks[0].Content.End, blocks[1].Content.Start);
        paragraphBreakRange.Delete();

        // Delete content of 2nd run
        blocks.Cast<Paragraph>(0).Inlines[1].Content.Delete();

        document.Save("Delete Content.docx");
    }
}

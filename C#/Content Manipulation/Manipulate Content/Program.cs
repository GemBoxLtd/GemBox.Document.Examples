using GemBox.Document;
using System;
using System.Linq;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("ManipulateContent.docx");
        var section = document.Sections[0];

        // Set content of 1st paragraph using plain text.
        section.Blocks[0].Content.LoadText("Inserted plain text to first paragraph.");

        // Set content of 2nd paragraph using hyperlink.
        var hyperlink = new Hyperlink(document, "https://www.gemboxsoftware.com/", "Inserted hyperlink.");
        section.Blocks[1].Content.Set(hyperlink.Content);

        // Insert HTML text at the end of 3rd paragraph.
        section.Blocks[2].Content.End
            .LoadText("<p style='color:orange'>Inserted HTML text with orange color.</p>",
                new HtmlLoadOptions() { InheritCharacterFormat = true, InheritParagraphFormat = true });

        // Insert picture at the beginning of last paragraph.
        var picture = new Picture(document, "Dices.png", 40, 30);
        section.Blocks.Last().Content.Start.InsertRange(picture.Content);

        document.Save("InsertContent.docx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("ManipulateContent.docx");
        var section = document.Sections[0];

        // Get content from 1st paragraph.
        ContentRange firstParagraphContent = section.Blocks[0].Content;
        Console.WriteLine(firstParagraphContent.ToString());

        // Get content from 2nd and 3rd paragraphs.
        ContentRange multipleParagraphsContent = new ContentRange(
            section.Blocks[1].Content.Start,
            section.Blocks[2].Content.End);
        Console.WriteLine(multipleParagraphsContent.ToString());
    }

    static void Example3()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("ManipulateContent.docx");
        var section = document.Sections[0];

        // Delete content from 1st and 2nd paragraph.
        ContentRange multipleParagraphsContent = new ContentRange(
            section.Blocks[0].Content.Start,
            section.Blocks[1].Content.End);
        multipleParagraphsContent.Delete();

        // Delete content from last (4th) paragraph.
        ContentRange lastParagraphContent = section.Blocks.Last().Content;
        lastParagraphContent.Delete();

        document.Save("DeleteContent.docx");
    }
}

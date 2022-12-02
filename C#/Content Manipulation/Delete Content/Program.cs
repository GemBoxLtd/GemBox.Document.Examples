using System.Linq;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");

        // Delete 1st paragraph's inlines.
        var paragraph1 = document.Sections[0].Blocks.Cast<Paragraph>(0);
        paragraph1.Inlines.Content.Delete();

        // Delete 3rd and 4th run from the 2nd paragraph.
        var paragraph2 = document.Sections[0].Blocks.Cast<Paragraph>(1);
        var runsContent = new ContentRange(
            paragraph2.Inlines[2].Content.Start,
            paragraph2.Inlines[3].Content.End);
        runsContent.Delete();

        // Delete specified text content.
        var bracketContent = document.Content.Find("(").First();
        bracketContent.Delete();

        document.Save("Delete Content.docx");
    }
}
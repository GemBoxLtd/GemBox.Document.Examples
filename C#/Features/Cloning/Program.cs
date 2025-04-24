using GemBox.Document;
using System.Linq;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");

        // Get first Section element.
        var section = document.Sections[0];

        // Get first Paragraph element.
        var paragraph = section.Blocks.OfType<Paragraph>().First();

        // Clone paragraph and add it to section.
        var cloneParagraph = paragraph.Clone(true);
        section.Blocks.Add(cloneParagraph);

        // Clone section and add it to document.
        var cloneSection = section.Clone(true);
        document.Sections.Add(cloneSection);

        document.Save("Cloning.docx");
    }
}

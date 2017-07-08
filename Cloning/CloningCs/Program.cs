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

        // Clone section
        document.Sections.Add(document.Sections[0].Clone(true));

        // Clone paragraphs
        foreach (Block item in document.Sections[0].Blocks)
            document.Sections.Last().Blocks.Add(item.Clone(true));

        document.Save("Cloning.docx");
    }
}

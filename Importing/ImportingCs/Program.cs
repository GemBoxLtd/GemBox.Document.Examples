using System;
using System.IO;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("Invoice.docx");

        string pathToFileDirectory = "Resources";

        DocumentModel sourceDocument = DocumentModel.Load(Path.Combine(pathToFileDirectory, "Reading.docx"), LoadOptions.DocxDefault);

        // Reuse same mapping for importing all sections to improve performance.
        var mapping = new ImportMapping(sourceDocument, document, false);

        // Import all sections from source document.
        foreach (Section sourceSection in sourceDocument.Sections)
        {
            Section destinationSection = document.Import(sourceSection, true, mapping);
            document.Sections.Add(destinationSection);
        }

        document.Save("Importing.docx");
    }
}

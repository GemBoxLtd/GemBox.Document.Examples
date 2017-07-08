using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        
        DocumentModel document = DocumentModel.Load("Reading.docx");

        // In order to achieve the conversion of a loaded Word file to PDF,
        // or to some other Word format,
        // we just need to save a DocumentModel object to desired output file format.

        document.Save("Convert.pdf");
    }
}

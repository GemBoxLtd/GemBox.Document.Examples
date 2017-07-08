using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("HtmlExport.docx");

        // Images will be embedded directly in HTML img src attribute.
        document.Save("Html Export.html", new HtmlSaveOptions() { EmbedImages = true });
    }
}

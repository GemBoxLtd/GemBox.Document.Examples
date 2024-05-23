using GemBox.Document;
using System;
using System.IO;
using System.IO.Compression;

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

        // In order to convert Word to PDF, we just need to:
        // 1. Load DOC or DOCX file into DocumentModel object.
        // 2. Save DocumentModel object to PDF file.
        DocumentModel document = DocumentModel.Load("Input.docx");
        document.Save("Output1.pdf");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load Word file.
        DocumentModel document = DocumentModel.Load("Input.docx");

        // Get Word pages.
        var pages = document.GetPaginator().Pages;

        // Create PDF save options.
        var pdfSaveOptions = new PdfSaveOptions() { ImageDpi = 220 };

        // Create ZIP file for storing PDF files.
        using (var archiveStream = File.OpenWrite("Output2.zip"))
        using (var archive = new ZipArchive(archiveStream, ZipArchiveMode.Create))
            // Iterate through Word pages.
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
            {
                DocumentModelPage page = pages[pageIndex];

                // Create ZIP entry for each document page.
                var entry = archive.CreateEntry($"Page {pageIndex + 1}.pdf");

                // Save each document page as PDF to ZIP entry.
                using (var pdfStream = new MemoryStream())
                using (var entryStream = entry.Open())
                {
                    page.Save(pdfStream, pdfSaveOptions);
                    pdfStream.CopyTo(entryStream);
                }
            }
    }

    static void Example3()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        PdfConformanceLevel conformanceLevel = PdfConformanceLevel.PdfA1a;

        // Load Word file.
        DocumentModel document = DocumentModel.Load("Input.docx");

        // Create PDF save options.
        var options = new PdfSaveOptions()
        {
            ConformanceLevel = conformanceLevel
        };

        // Save to PDF file.
        document.Save("Output3.pdf", options);
    }
}

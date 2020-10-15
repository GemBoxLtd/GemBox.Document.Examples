using System;
using System.IO;
using System.IO.Compression;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // Load a Word file into the DocumentModel object.
        var document = DocumentModel.Load("Input.docx");
        var section = document.Sections[0];

        // Calculate the default image width based on the page size and DPI resolution.
        var widthInPoints = section.PageSetup.PageWidth;
        var resolution = 300;
        var widthInPixels = widthInPoints * resolution / 72;

        var imageOptions = new ImageSaveOptions(ImageSaveFormat.Png)
        {
            // Select the first Word page.
            PageNumber = 0,

            // Set the DPI resolution.
            DpiX = resolution,
            DpiY = resolution,

            // Set the image width to half and keep the aspect ratio.
            Width = widthInPixels / 2
        };

        // Save the DocumentModel object to a PNG file.
        document.Save("Output.png", imageOptions);
    }

    static void Example2()
    {
        // Load a Word file.
        var document = DocumentModel.Load("Input.docx");

        var imageOptions = new ImageSaveOptions(ImageSaveFormat.Tiff)
        {
            // Max integer value indicates that all document pages should be saved.
            PageCount = int.MaxValue
        };

        // Save the TIFF file with multiple frames, each frame represents a single Word page.
        document.Save("Output.tiff", imageOptions);
    }

    static void Example3()
    {
        // Load a Word file.
        var document = DocumentModel.Load("Input.docx");

        var imageOptions = new ImageSaveOptions(ImageSaveFormat.Png);

        // Get Word pages.
        var pages = document.GetPaginator().Pages;

        // Create a ZIP file for storing PNG files.
        using (var archiveStream = File.OpenWrite("Output.zip"))
        using (var archive = new ZipArchive(archiveStream, ZipArchiveMode.Create))
        {
            // Iterate through the Word pages.
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
            {
                DocumentModelPage page = pages[pageIndex];

                // Create a ZIP entry for each document page.
                var entry = archive.CreateEntry($"Page {pageIndex + 1}.png");

                // Save each document page as a PNG image to the ZIP entry.
                using (var imageStream = new MemoryStream())
                using (var entryStream = entry.Open())
                {
                    page.Save(imageStream, imageOptions);
                    imageStream.CopyTo(entryStream);
                }
            }
        }
    }
}

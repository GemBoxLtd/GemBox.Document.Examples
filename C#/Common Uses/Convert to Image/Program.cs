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

        // Create image save options.
        var imageOptions = new ImageSaveOptions(ImageSaveFormat.Png)
        {
            PageNumber = 0, // Select the first Word page.
            Width = 1240 // Set the image width and keep the aspect ratio.
        };

        // Save the DocumentModel object to a PNG file.
        document.Save("Output.png", imageOptions);
    }

    static void Example2()
    {
        // Load a Word file.
        var document = DocumentModel.Load("Input.docx");

        // Max integer value indicates that all document pages should be saved.
        var imageOptions = new ImageSaveOptions(ImageSaveFormat.Tiff)
        {
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

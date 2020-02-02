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
        Example4();
    }

    static void Example1()
    {
        // In order to convert Word to PDF, we just need to:
        // 1. Load DOC or DOCX file into DocumentModel object.
        // 2. Save DocumentModel object to PDF file.
        DocumentModel document = DocumentModel.Load("Input.docx");
        document.Save("Output1.pdf");
    }

    static void Example2()
    {
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
        // Load input HTML file.
        DocumentModel document = DocumentModel.Load("Input.html");

        // When reading any HTML content a single Section element is created.
        // We can use that Section element to specify various page options.
        Section section = document.Sections[0];
        PageSetup pageSetup = section.PageSetup;
        PageMargins pageMargins = pageSetup.PageMargins;
        pageMargins.Top = pageMargins.Bottom = pageMargins.Left = pageMargins.Right = 0;

        // Save output PDF file.
        document.Save("Output3.pdf");
    }

    static void Example4()
    {
        var html = @"
<html>
<style>
  @page {
    size: A5 landscape;
    margin: 6cm 1cm 1cm;
    mso-header-margin: 1cm;
    mso-footer-margin: 1cm;
  }

  body {
    background: #EDEDED;
    border: 1pt solid black;
    padding: 20pt;
  }

  br {
    page-break-before: always;
  }

  p { margin: 0; }
  header { color: #FF0000; text-align: center; }
  main { color: #00B050; }
  footer { color: #0070C0; text-align: right; }
</style>

<body>
  <header>
    <p>Header text.</p>
  </header>
  <main>
    <p>First page.</p>
    <br>
    <p>Second page.</p>
    <br>
    <p>Third page.</p>
    <br>
    <p>Fourth page.</p>
  </main>
  <footer>
    <p>Footer text.</p>
    <p>Page <span style='mso-field-code:PAGE'>1</span> of <span style='mso-field-code:NUMPAGES'>1</span></p>
  </footer>
</body>
</html>";

        var htmlLoadOptions = new HtmlLoadOptions();
        using (var htmlStream = new MemoryStream(htmlLoadOptions.Encoding.GetBytes(html)))
        {
            // Load input HTML text as stream.
            var document = DocumentModel.Load(htmlStream, htmlLoadOptions);
            // Save output PDF file.
            document.Save("Output4.pdf");
        }
    }
}

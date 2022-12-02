using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        // Load Word file (DOC, DOCX, RTF, XML) into DocumentModel object.
        var document = DocumentModel.Load("ExportToHtml.docx");

        var saveOptions = new HtmlSaveOptions()
        {
            HtmlType = HtmlType.Html,
            EmbedImages = true,
            UseSemanticElements = true
        };

        // Save DocumentModel object to HTML (or MHTML) file.
        document.Save("Exported.html", saveOptions);
    }

    static void Example2()
    {
        // Load input HTML file.
        DocumentModel document = DocumentModel.Load("Input.html");

        // When reading any HTML content a single Section element is created,
        // which can be used to specify various Word document's page options.
        // The same can also be achieved with HTML document itself,
        // by using CSS properties on "@page" directive or "<body>" element.
        Section section = document.Sections[0];
        PageSetup pageSetup = section.PageSetup;
        PageMargins pageMargins = pageSetup.PageMargins;
        pageMargins.Top = pageMargins.Bottom = pageMargins.Left = pageMargins.Right = 0;

        // Save output DOCX file.
        document.Save("Output.docx");
    }
}

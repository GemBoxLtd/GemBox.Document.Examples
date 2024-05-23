using DocumentCorePages.Models;
using GemBox.Document;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System;
using System.Collections.Generic;
using System.IO;

namespace DocumentCorePages.Pages
{
    public class IndexModel : PageModel
    {
        private readonly IWebHostEnvironment environment;

        [BindProperty]
        public InvoiceModel Invoice { get; set; }

        // If using the Professional version, put your serial key below.
        static IndexModel() => ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        public IndexModel(IWebHostEnvironment environment)
        {
            this.environment = environment;
            this.Invoice = new InvoiceModel();
        }

        public void OnGet() { }

        public FileContentResult OnPost()
        {
            // Load template document.
            var path = Path.Combine(this.environment.ContentRootPath, "InvoiceWithPlaceholders.docx");
            var document = DocumentModel.Load(path);

            // Execute find and replace operations.
            document.Content.Replace("{{Number}}", this.Invoice.Number.ToString("0000"));
            document.Content.Replace("{{Date}}", this.Invoice.Date.ToString("d MMM yyyy HH:mm"));
            document.Content.Replace("{{Company}}", this.Invoice.Company);
            document.Content.Replace("{{Address}}", this.Invoice.Address);
            document.Content.Replace("{{Name}}", this.Invoice.Name);

            // Save document in specified file format.
            using var stream = new MemoryStream();
            document.Save(stream, this.Invoice.Options);

            // Download file.
            return File(stream.ToArray(), this.Invoice.Options.ContentType, $"OutputFromPage.{this.Invoice.Format.ToLower()}");
        }
    }
}

namespace DocumentCorePages.Models
{
    public class InvoiceModel
    {
        public int Number { get; set; } = 1;
        public DateTime Date { get; set; } = DateTime.Today;
        public string Company { get; set; } = "ACME Corp.";
        public string Address { get; set; } = "240 Old Country Road, Springfield, United States";
        public string Name { get; set; } = "Joe Smith";
        public string Format { get; set; } = "PDF";
        public SaveOptions Options => this.FormatMappingDictionary[this.Format];
        public IDictionary<string, SaveOptions> FormatMappingDictionary => new Dictionary<string, SaveOptions>()
        {
            ["PDF"] = new PdfSaveOptions(),
            ["DOCX"] = new DocxSaveOptions(),
            ["ODT"] = new OdtSaveOptions(),
            ["HTML"] = new HtmlSaveOptions() { EmbedImages = true },
            ["MHTML"] = new HtmlSaveOptions() { HtmlType = HtmlType.Mhtml },
            ["RTF"] = new RtfSaveOptions(),
            ["XML"] = new XmlSaveOptions(),
            ["TXT"] = new TxtSaveOptions(),
            ["XPS"] = new XpsSaveOptions(), // XPS is supported only on Windows.
            ["PNG"] = new ImageSaveOptions(ImageSaveFormat.Png),
            ["JPG"] = new ImageSaveOptions(ImageSaveFormat.Jpeg),
            ["BMP"] = new ImageSaveOptions(ImageSaveFormat.Bmp),
            ["GIF"] = new ImageSaveOptions(ImageSaveFormat.Gif),
            ["TIF"] = new ImageSaveOptions(ImageSaveFormat.Tiff),
            ["SVG"] = new ImageSaveOptions(ImageSaveFormat.Svg)
        };
    }
}

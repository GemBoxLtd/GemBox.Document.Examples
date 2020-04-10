using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using GemBox.Document;

namespace DocumentCore.Controllers
{
    public class DocumentController : Controller
    {
        private static readonly SelectListItem[] Countries = CultureInfo.GetCultures(CultureTypes.SpecificCultures)
            .Select(c => new RegionInfo(c.LCID).EnglishName)
            .Distinct()
            .OrderBy(k => k)
            .Select(k => new SelectListItem() { Text = k, Value = k })
            .ToArray();

        private readonly IWebHostEnvironment environment;

        public DocumentController(IWebHostEnvironment environment)
        {
            this.environment = environment;
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        [HttpGet]
        public IActionResult Create()
        {
            return View(new InvoiceModel()
            {
                Number = 1,
                Date = DateTime.Today,
                Company = "ACME Corp.",
                Address = "240 Old Country Road, Springfield, IL",
                Countries = Countries,
                Country = "United States",
                FullName = "Joe Smith",
                SelectedFormat = "DOCX"
            });
        }

        [HttpPost]
        public ActionResult Create(InvoiceModel model)
        {
            if (!ModelState.IsValid)
                return View(model);

            SaveOptions options = GetSaveOptions(model.SelectedFormat);
            DocumentModel document = this.Process(model);

            using (var stream = new MemoryStream())
            {
                document.Save(stream, options);
                return File(stream.ToArray(), options.ContentType, "Create." + model.SelectedFormat.ToLower());
            }
        }

        private static SaveOptions GetSaveOptions(string format)
        {
            switch (format.ToUpper())
            {
                case "DOCX":
                    return SaveOptions.DocxDefault;
                case "HTML":
                    return SaveOptions.HtmlDefault;
                case "RTF":
                    return SaveOptions.RtfDefault;
                case "XML":
                    return SaveOptions.XmlDefault;
                case "TXT":
                    return SaveOptions.TxtDefault;
                case "PDF":
                    return SaveOptions.PdfDefault;

                case "XPS":
                case "PNG":
                case "JPG":
                case "GIF":
                case "TIF":
                case "BMP":
                case "WMP":
                    throw new InvalidOperationException("To enable saving to XPS or image format, add 'Microsoft.WindowsDesktop.App' framework reference.");

                default:
                    throw new NotSupportedException();
            }
        }

        private DocumentModel Process(InvoiceModel model)
        {
            string path = Path.Combine(this.environment.ContentRootPath, "Invoice.docx");

            // Load template document.
            DocumentModel document = DocumentModel.Load(path);

            // Execute mail merge process.
            document.MailMerge.Execute(model);

            return document;
        }
    }

    public class InvoiceModel
    {
        public int Number { get; set; }
        public DateTime Date { get; set; }
        public string Company { get; set; }
        public string Address { get; set; }
        public IList<SelectListItem> Countries { get; set; }
        public string Country { get; set; }
        public string SelectedFormat { get; set; }
        public string FullName { get; set; }
    }
}
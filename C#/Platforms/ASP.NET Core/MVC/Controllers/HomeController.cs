﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using DocumentCoreMvc.Models;
using GemBox.Document;

namespace DocumentCoreMvc.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment environment;

        public HomeController(IWebHostEnvironment environment)
        {
            this.environment = environment;

            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        public IActionResult Index()
        {
            return View(new InvoiceModel());
        }

        public FileStreamResult Download(InvoiceModel model)
        {
            // Load template document.
            var path = Path.Combine(this.environment.ContentRootPath, "InvoiceWithFields.docx");
            var document = DocumentModel.Load(path);

            // Execute mail merge process.
            document.MailMerge.Execute(model);

            // Save document in specified file format.
            var stream = new MemoryStream();
            document.Save(stream, model.Options);

            // Download file.
            return File(stream, model.Options.ContentType, $"OutputFromView.{model.Format.ToLower()}");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel() { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}

namespace DocumentCoreMvc.Models
{
    public class InvoiceModel
    {
        public int Number { get; set; } = 1;
        public DateTime Date { get; set; } = DateTime.Today;
        public string Company { get; set; } = "ACME Corp.";
        public string Address { get; set; } = "240 Old Country Road, Springfield, United States";
        public string Name { get; set; } = "Joe Smith";
        public string Format { get; set; } = "DOCX";
        public SaveOptions Options => this.FormatMappingDictionary[this.Format];
        public IDictionary<string, SaveOptions> FormatMappingDictionary => new Dictionary<string, SaveOptions>()
        {
            ["DOCX"] = new DocxSaveOptions(),
            ["HTML"] = new HtmlSaveOptions() { EmbedImages = true },
            ["RTF"] = new RtfSaveOptions(),
            ["TXT"] = new TxtSaveOptions(),
            ["PDF"] = new PdfSaveOptions(),
            ["XPS"] = new XpsSaveOptions(),
            ["XML"] = new XmlSaveOptions(),
            ["BMP"] = new ImageSaveOptions(ImageSaveFormat.Bmp),
            ["PNG"] = new ImageSaveOptions(ImageSaveFormat.Png),
            ["JPG"] = new ImageSaveOptions(ImageSaveFormat.Jpeg),
            ["GIF"] = new ImageSaveOptions(ImageSaveFormat.Gif),
            ["TIF"] = new ImageSaveOptions(ImageSaveFormat.Tiff)
        };
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Linq;
using GemBox.Document;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace Document.Samples.Core.Controllers
{
    public class DocumentController : Controller
    {
        private static readonly IList<InvoiceItemModel> data = new List<InvoiceItemModel>()
        {
            new InvoiceItemModel() { Date = DateTime.UtcNow.AddDays(-4), Hours = 8, Price = 35.0 },
            new InvoiceItemModel() { Date = DateTime.UtcNow.AddDays(-3), Hours = 8, Price = 35.0 },
            new InvoiceItemModel() { Date = DateTime.UtcNow.AddDays(-2), Hours = 6, Price = 35.0 },
            new InvoiceItemModel() { Date = DateTime.UtcNow.AddDays(-1), Hours = 5, Price = 35.0 },
        };

        private static readonly SelectListItem[] countries = CultureInfo.GetCultures(CultureTypes.SpecificCultures)
            .Select(c => new RegionInfo(c.LCID).EnglishName)
            .Distinct()
            .OrderBy(k => k)
            .Select(k => new SelectListItem() { Text = k, Value = k })
            .ToArray();

        private static int invoiceNumber = 1;
        private IHostingEnvironment environment;

        public DocumentController(IHostingEnvironment environment)
        {
            this.environment = environment;
        }

        private static SaveOptions GetSaveOptions(string format)
        {
            switch (format.ToUpperInvariant())
            {
                case "DOCX":
                    return SaveOptions.DocxDefault;
                case "HTML":
                    return SaveOptions.HtmlDefault;
                case "RTF":
                    return SaveOptions.RtfDefault;
                case "TXT":
                    return SaveOptions.TxtDefault;
                default:
                    throw new NotSupportedException("Format '" + format + "' is not supported.");
            }
        }

        private static byte[] GetBytes(DocumentModel document, SaveOptions options)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                document.Save(stream, options);
                return stream.ToArray();
            }
        }

        private DocumentModel Process(InvoiceModel model)
        {
            string path = Path.Combine(this.environment.ContentRootPath, "Invoice.docx");

            // Load template document
            DocumentModel document = DocumentModel.Load(path);

            // Subscribe to FieldMerging event (we want to format the output)
            document.MailMerge.FieldMerging += (sender, e) =>
            {
                if (e.IsValueFound)
                    switch (e.FieldName)
                    {
                        case "Date":
                            ((Run)e.Inline).Text = ((DateTime)e.Value).ToString("dddd, MMMM d, yyyy", CultureInfo.InvariantCulture);
                            break;
                        case "Price":
                        case "Total":
                        case "GrandTotal":
                            ((Run)e.Inline).Text = InvoiceModel.FormatNumber((double)e.Value);
                            break;
                    }
            };

            // Fill table and grand total
            document.MailMerge.Execute(model.Items, "Item");
            document.MailMerge.Execute(model);

            // Fill invoice data
            document.MailMerge.Execute(
                new
                {
                    Number = model.Number,
                    InvoiceDate = model.Date.ToShortDateString()
                });


            // Fill customer data
            document.MailMerge.Execute(
                new Dictionary<string, object>()
                {
                    { "Name", model.Name },
                    { "Address" , model.Address },
                    { "Country" , model.Country },
                    { "ContactPerson" , model.ContactPerson }
                });

            // Fill notes
            document.MailMerge.Execute(
                new KeyValuePair<string, object>[]
                {
                    new KeyValuePair<string, object>("Notes", model.Notes)
                });


            invoiceNumber = invoiceNumber == int.MaxValue ? 1 : invoiceNumber + 1;

            return document;
        }

        public IActionResult Create()
        {
            return View(new InvoiceModel()
            {
                Number = invoiceNumber,
                Date = DateTime.UtcNow,
                Name = "ACME Corp.",
                Address = "240 Old Country Road, Springfield, IL",
                Countries = countries,
                Country = "United States",
                ContactPerson = "Joe Smith",
                Items = data,
                Notes = "Payment via check.",
                SelectedFormat = "DOCX"
            });
        }

        [HttpPost]
        public ActionResult Create(InvoiceModel model)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            if (!ModelState.IsValid)
                return View(model);

            SaveOptions options = GetSaveOptions(model.SelectedFormat);
            DocumentModel document = this.Process(model);

            return File(GetBytes(document, options), options.ContentType, "Create." + model.SelectedFormat.ToLowerInvariant());
        }
    }

    public class InvoiceModel
    {
        public int Number { get; set; }
        public DateTime Date { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public IList<SelectListItem> Countries { get; set; }
        public string Country { get; set; }
        public IList<InvoiceItemModel> Items { get; set; }
        public string Notes { get; set; }
        public string SelectedFormat { get; set; }

        [Display(Name = "Contact person")]
        public string ContactPerson { get; set; }

        public static string FormatNumber(double value)
        {
            return value.ToString("F2", CultureInfo.InvariantCulture);
        }

        public double GrandTotal
        {
            get { return this.Items.Select(i => i.Total).Sum(); }
        }

        public string GrandTotalText
        {
            get { return FormatNumber(this.GrandTotal); }
        }
    }

    public class InvoiceItemModel
    {
        public DateTime Date { get; set; }
        public int Hours { get; set; }
        public double Price { get; set; }

        public string PriceText
        {
            get { return InvoiceModel.FormatNumber(this.Price); }
            set { this.Price = double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double result) ? result : 0.0; }
        }

        public double Total
        {
            get { return this.Hours * this.Price; }
        }

        public string TotalText
        {
            get { return InvoiceModel.FormatNumber(this.Hours * this.Price); }
        }
    }
}
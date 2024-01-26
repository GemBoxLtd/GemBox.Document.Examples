using System.IO;
using System.Web.Mvc;
using AspNetWordEditor.Models;
using GemBox.Document;

namespace AspNetWordEditor.Controllers
{
    public class HomeController : Controller
    {
        [HttpGet]
        public ActionResult Index() => this.View(new FileModel());

        [HttpPost]
        public FileResult Download(FileModel model)
        {
            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var templateFile = Server.MapPath("~/App_Data/DocumentTemplate.docx");

            // Load template document.
            var document = DocumentModel.Load(templateFile);

            // Insert content from HTML editor.
            var bookmark = document.Bookmarks["HtmlBookmark"];
            bookmark.GetContent(true).LoadText(model.Content, LoadOptions.HtmlDefault);

            // Save document to stream in specified format.
            var saveOptions = GetSaveOptions(model.Extension);
            var stream = new MemoryStream();
            document.Save(stream, saveOptions);

            // Download document.
            var downloadFile = $"Output{model.Extension}";
            return File(stream.ToArray(), saveOptions.ContentType, downloadFile);
        }

        private static SaveOptions GetSaveOptions(string extension)
        {
            switch (extension)
            {
                case ".pdf": return SaveOptions.PdfDefault;
                case ".docx": return SaveOptions.DocxDefault;
                case ".odt": return SaveOptions.OdtDefault;
                case ".html": return SaveOptions.HtmlDefault;
                case ".mhtml": return new HtmlSaveOptions() { HtmlType = HtmlType.Mhtml };
                case ".rtf": return SaveOptions.RtfDefault;
                case ".xml": return SaveOptions.XmlDefault;
                case ".xps": return SaveOptions.XpsDefault;
                case ".png": return SaveOptions.ImageDefault;
                case ".jpeg": return new ImageSaveOptions(ImageSaveFormat.Jpeg);
                case ".bmp": return new ImageSaveOptions(ImageSaveFormat.Bmp);
                case ".gif": return new ImageSaveOptions(ImageSaveFormat.Gif);
                case ".tiff": return new ImageSaveOptions(ImageSaveFormat.Tiff);
                case ".wmp": return new ImageSaveOptions(ImageSaveFormat.Wmp);
                case ".svg": return new ImageSaveOptions(ImageSaveFormat.Svg);
                default: return SaveOptions.TxtDefault;
            }
        }
    }
}
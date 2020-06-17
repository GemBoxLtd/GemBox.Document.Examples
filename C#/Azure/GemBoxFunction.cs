using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using GemBox.Document;

public static class GemBoxFunction
{
    [FunctionName("GemBoxFunction")]
    public static async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req,
        ILogger log)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        Section section = new Section(document);
        document.Sections.Add(section);

        Paragraph paragraph = new Paragraph(document);
        section.Blocks.Add(paragraph);

        Run run = new Run(document, "Hello World!");
        paragraph.Inlines.Add(run);

        var fileName = "Output.docx";
        var options = SaveOptions.DocxDefault;

        using (var stream = new MemoryStream())
        {
            document.Save(stream, options);
            return new FileContentResult(stream.ToArray(), options.ContentType) { FileDownloadName = fileName };
        }
    }
}
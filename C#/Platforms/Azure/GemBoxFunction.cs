using System.IO;
using System.Net;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using GemBox.Document;

public class GemBoxFunction
{
    [Function("GemBoxFunction")]
    public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequestData req)
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        var paragraph = new Paragraph(document);
        section.Blocks.Add(paragraph);

        var run = new Run(document, "Hello World!");
        paragraph.Inlines.Add(run);

        string fileName = "Output.docx";
        var options = SaveOptions.DocxDefault;

        using var stream = new MemoryStream();
        document.Save(stream, options);
        var bytes = stream.ToArray();

        var response = req.CreateResponse(HttpStatusCode.OK);
        response.Headers.Add("Content-Type", options.ContentType);
        response.Headers.Add("Content-Disposition", "attachment; filename=" + fileName);
        await response.Body.WriteAsync(bytes, 0, bytes.Length);
        return response;
    }
}
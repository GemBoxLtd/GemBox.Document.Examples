Imports System.IO
Imports System.Net
Imports System.Threading.Tasks
Imports Microsoft.Azure.Functions.Worker
Imports Microsoft.Azure.Functions.Worker.Http
Imports GemBox.Document

Public Class GemBoxFunction
    <[Function]("GemBoxFunction")>
    Public Async Function Run(<HttpTrigger(AuthorizationLevel.Anonymous, "get")> req As HttpRequestData) As Task(Of HttpResponseData)

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        Dim paragraph As New Paragraph(document)
        section.Blocks.Add(paragraph)

        Dim inline As New Run(document, "Hello World!")
        paragraph.Inlines.Add(inline)

        Dim fileName = "Output.docx"
        Dim options = SaveOptions.DocxDefault

        Using stream As New MemoryStream()
            document.Save(stream, options)
            Dim bytes = stream.ToArray()

            Dim response = req.CreateResponse(HttpStatusCode.OK)
            response.Headers.Add("Content-Type", options.ContentType)
            response.Headers.Add("Content-Disposition", "attachment; filename=" & fileName)
            Await response.Body.WriteAsync(bytes, 0, bytes.Length)
            Return response
        End Using

    End Function
End Class

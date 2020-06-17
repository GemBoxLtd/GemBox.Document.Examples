Imports System.IO
Imports Microsoft.AspNetCore.Mvc
Imports Microsoft.Azure.WebJobs
Imports Microsoft.Azure.WebJobs.Extensions.Http
Imports Microsoft.AspNetCore.Http
Imports Microsoft.Extensions.Logging
Imports GemBox.Document

Module GemBoxFunction
        <FunctionName("GemBoxFunction")>
        Async Function Run(
            <HttpTrigger(AuthorizationLevel.Anonymous, "get", Route := Nothing)> req As HttpRequest,
            log As ILogger) as Task(Of IActionResult)
        
            ' If using Professional version, put your serial key below.
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

            Using stream As new MemoryStream()
                document.Save(stream, options)
                return New FileContentResult(stream.ToArray(), options.ContentType) With { .FileDownloadName = fileName } 
            End Using
             
        End Function
End Module

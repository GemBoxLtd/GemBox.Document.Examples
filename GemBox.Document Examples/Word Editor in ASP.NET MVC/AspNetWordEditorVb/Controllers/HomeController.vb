Imports System.IO
Imports System.Web.Mvc
Imports AspNetWordEditor.Models
Imports GemBox.Document

Namespace Controllers
    Public Class HomeController
        Inherits Controller

        <HttpGet>
        Function Index() As ActionResult
            Return View()
        End Function

        <HttpPost>
        Function Download(model As FileModel) As FileResult

            ' If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY")

            Dim templateFile = Server.MapPath("~/App_Data/%#DocumentTemplate.docx%")

            ' Load template document.
            Dim document = DocumentModel.Load(templateFile)

            ' Insert content from HTML editor.
            Dim bookmark = document.Bookmarks("HtmlBookmark")
            bookmark.GetContent(True).LoadText(model.Content, LoadOptions.HtmlDefault)

            ' Save document to stream in specified format.
            Dim saveOptions = GetSaveOptions(model.Extension)
            Dim stream As New MemoryStream()
            document.Save(stream, saveOptions)

            ' Download document.
            Dim downloadFile = $"Output{model.Extension}"
            Return File(stream.ToArray(), saveOptions.ContentType, downloadFile)

        End Function

        Private Shared Function GetSaveOptions(extension As String) As SaveOptions
            Select Case extension
                Case ".docx"
                    Return SaveOptions.DocxDefault
                Case ".pdf"
                    Return SaveOptions.PdfDefault
                Case ".xps"
                    Return SaveOptions.XpsDefault
                Case ".html"
                    Return SaveOptions.HtmlDefault
                Case ".mhtml"
                    Return New HtmlSaveOptions() With {.HtmlType = HtmlType.Mhtml}
                Case ".rtf"
                    Return SaveOptions.RtfDefault
                Case ".xml"
                    Return SaveOptions.XmlDefault
                Case ".png"
                    Return SaveOptions.ImageDefault
                Case ".jpeg"
                    Return New ImageSaveOptions(ImageSaveFormat.Jpeg)
                Case ".gif"
                    Return New ImageSaveOptions(ImageSaveFormat.Gif)
                Case ".bmp"
                    Return New ImageSaveOptions(ImageSaveFormat.Bmp)
                Case ".tiff"
                    Return New ImageSaveOptions(ImageSaveFormat.Tiff)
                Case ".wmp"
                    Return New ImageSaveOptions(ImageSaveFormat.Wmp)
                Case Else
                    Return SaveOptions.TxtDefault
            End Select
        End Function

    End Class
End Namespace
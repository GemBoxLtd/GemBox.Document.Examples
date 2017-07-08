Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("HtmlExport.docx")

        ' Images will be embedded directly in HTML img src attribute.
        document.Save("Html Export.html", New HtmlSaveOptions() With {.EmbedImages = True})

    End Sub

End Module
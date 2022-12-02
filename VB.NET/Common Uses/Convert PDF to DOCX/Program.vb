Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("CustomInvoice.pdf",
            New PdfLoadOptions() With
            {
                .LoadType = PdfLoadType.HighFidelity
            })

        document.Save("ConvertedFromPdf.docx")
        
    End Sub
End Module
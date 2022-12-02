Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("MergeBarcodes.docx")

        ' Create data source for mail merge process.
        Dim data = New With
        {
            .QrCode = "QR Code created with GemBox.Document",
            .Code128 = "1234567890",
            .Ean13 = "5901234123457"
        }

        ' Execute mail merge process.
        document.MailMerge.Execute(data)

        document.Save("Mail Merge Output.docx")

    End Sub
End Module
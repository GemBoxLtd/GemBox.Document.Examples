Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Create a new document.
        Dim document As New DocumentModel()
        document.DefaultParagraphFormat.Alignment = HorizontalAlignment.Center

        ' Create a barcode merge field that will display the value in a barcode font.
        Dim barcodeField As New Field(document, FieldType.MergeField, "Barcode", "«Barcode»") With
        {
            .CharacterFormat = New CharacterFormat() With
            {
                .FontName = "Code 128",
                .Size = 80
            }
        }

        ' Create a label merge field that will display the value with a '*' character as the prefix and suffix.
        Dim labelField As New Field(document, FieldType.MergeField, "Label \b * \f *", "*«Label»*") With
        {
            .CharacterFormat = New CharacterFormat() With
            {
                .FontName = "Arial Black",
                .Size = 20,
                .FontColor = Color.Red
            }
        }

        ' Add merge fields to the document.
        document.Sections.Add(
            New Section(document,
                New Paragraph(document, barcodeField),
                New Paragraph(document, labelField)))

        document.MailMerge.Execute(New With {.Barcode = "1234567890", .Label = "1234567890"})

        document.Save("Barcode Output.docx")

    End Sub
End Module
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        ' Add '{ AUTHOR }' field.
        section.Blocks.Add(
            New Paragraph(document,
                New Run(document, "Author: "),
                New Field(document, FieldType.Author, Nothing, "Mario at GemBox")))

        ' Add '{ DATE }' field.
        section.Blocks.Add(
            New Paragraph(document,
                New Run(document, "Date: "),
                New Field(document, FieldType.Date)))

        ' Add '{ DATE \@ "dddd, MMMM dd, yyyy" }' field.
        section.Blocks.Add(
            New Paragraph(document,
                New Run(document, "Date with specified format: "),
                New Field(document, FieldType.Date, "\@ ""dddd, MMMM dd, yyyy""")))

        document.Save("Fields.docx")

    End Sub
End Module

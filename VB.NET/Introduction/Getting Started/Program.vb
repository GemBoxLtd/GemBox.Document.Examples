Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        Dim paragraph As New Paragraph(document)
        section.Blocks.Add(paragraph)

        Dim run As New Run(document, "Hello World!")
        paragraph.Inlines.Add(run)

        document.Save("HelloWorld.docx")

    End Sub

End Module

Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, " ")))

        Dim viewOptions = document.ViewOptions
        viewOptions.ViewType = ViewType.Web
        viewOptions.Zoom = 75

        document.Save("View Options.docx")

    End Sub
End Module
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "This document has been opened in read-only mode."),
                New Paragraph(document, "Changes cannot be made to the original document."),
                New Paragraph(document, "To save changes a new copy of the document must be created.")))

        Dim protection As WriteProtection = document.WriteProtection
        protection.SetPassword("pass")

        document.Save("DOCX Write Protection.docx")

    End Sub
End Module
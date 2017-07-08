Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        Dim section = New Section(document)
        document.Sections.Add(section)

        Dim paragraph1 = New Paragraph(document, "This document has been opened in read-only mode. Changes cannot be made to the original document. To save changes, create a new copy of the document.")
        section.Blocks.Add(paragraph1)

        Dim paragraph2 = New Paragraph(document, "To enable modifying use password: 1234")
        section.Blocks.Add(paragraph2)

        Dim protection = document.WriteProtection
        ' For DOCX file format: disallow resaving the document using the same file name.
        protection.SetPassword("1234")

        document.Save("DOCX Write Protection.docx")

    End Sub

End Module
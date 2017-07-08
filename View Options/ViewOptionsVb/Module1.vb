Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "GemBox.Document is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way.")))

        document.ViewOptions.ViewType = ViewType.Print
        document.ViewOptions.Zoom = 75

        document.Save("View Options.docx")

    End Sub

End Module
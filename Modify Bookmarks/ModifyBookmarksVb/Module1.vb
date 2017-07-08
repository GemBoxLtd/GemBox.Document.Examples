Imports System
Imports System.IO
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("BookmarksTemplate.docx")

        Dim pathToResources As String = "Resources"

        document.Bookmarks("CompanyName").GetContent(False).LoadText("ACME Corp")
        document.Bookmarks("CompanyAddress").GetContent(False).LoadText("240 Old Country Road, Springfield, IL")
        document.Bookmarks("Country").GetContent(False).LoadText("USA")
        document.Bookmarks("ContactPerson").GetContent(False).LoadText("Joe Smith")

        document.Bookmarks("Text").GetContent(False).LoadText(
            "GemBox.Document is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way.",
            New CharacterFormat() With {.Size = 14})

        Dim picture = New Picture(document, Path.Combine(pathToResources, "Acme.png"))
        document.Bookmarks("Logo").GetContent(False).Set(picture.Content)

        document.Save("Modify Bookmarks.docx")

    End Sub

End Module
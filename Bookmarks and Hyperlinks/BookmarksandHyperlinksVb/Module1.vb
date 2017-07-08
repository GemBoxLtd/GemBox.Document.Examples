Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        Dim bookmarkName As String = "LinkToTop"

        ' When user clicks on this link it will jump to the text between BookmarkStart and BookmarkEnd. 
        document.Sections.Add(
            New Section(document,
                New Paragraph(document,
                    New BookmarkStart(document, bookmarkName),
                    New Run(document, "GemBox.Document"),
                    New BookmarkEnd(document, bookmarkName)),
                New Paragraph(document,
                    New Hyperlink(document, "http://www.gemboxsoftware.com/document/overview", "GemBox.Document"),
                    New Run(document, " is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way.")),
                New Paragraph(document),
                New Paragraph(document),
                New Paragraph(document),
             New Paragraph(document,
                 New Hyperlink(document, bookmarkName, "To Top") With {
                    .IsBookmarkLink = True
                })))

        document.Save("Bookmarks and Hyperlinks.docx")

    End Sub

End Module
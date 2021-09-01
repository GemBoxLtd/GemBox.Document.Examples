Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim bookmarkName = "TopOfDocument"

        document.Sections.Add(
            New Section(document,
                New Paragraph(document,
                    New BookmarkStart(document, bookmarkName),
                    New Run(document, "This is a 'TopOfDocument' bookmark."),
                    New BookmarkEnd(document, bookmarkName)),
                New Paragraph(document,
                    New Run(document, "The following is a link to "),
                    New Hyperlink(document, "https://www.gemboxsoftware.com/document", "GemBox.Document Overview"),
                    New Run(document, " page.")),
                 New Paragraph(document,
                    New SpecialCharacter(document, SpecialCharacterType.PageBreak),
                    New Run(document, "This is a document's second page."),
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Hyperlink(document, bookmarkName, "Return to 'TopOfDocument'.") With {.IsBookmarkLink = True})))

        document.Save("Bookmarks and Hyperlinks.docx")

    End Sub
End Module
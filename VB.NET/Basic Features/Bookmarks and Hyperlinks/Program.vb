Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()
        Example1()
        Example2()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
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

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("BookmarksTemplate.docx")

        ' Replace bookmark's content with plain text.
        document.Bookmarks("Company").GetContent(False).LoadText("Acme Corporation")

        ' Replace bookmark's content with HTML text.
        document.Bookmarks("Address").GetContent(False).LoadText(
            "<span style='font: italic 8pt Calibri; color: red;'>240 Old Country Road, Springfield, IL</span>",
            LoadOptions.HtmlDefault)

        ' Insert hyperlink into a bookmark.
        Dim hyperlink As New Hyperlink(document, "mailto:joe.doe@acme.co", "joe.doe@acme.co")
        document.Bookmarks("Contact").GetContent(False).Set(hyperlink.Content)

        ' Insert image into a bookmark.
        Dim picture As New Picture(document, "Acme.png")
        document.Bookmarks("Logo").GetContent(False).Set(picture.Content)

        ' Insert text and table into a bookmark.
        Dim itemsRange As ContentRange = document.Bookmarks("Items").GetContent(False)
        itemsRange = itemsRange.LoadText("Storage:")
        Dim table As New Table(document, 6, 3, Function(r, c) New TableCell(document, New Paragraph(document, $"Item {r}-{c}")))
        itemsRange.End.InsertRange(table.Content)

        document.Save("Modified Bookmarks.docx")
    End Sub

End Module

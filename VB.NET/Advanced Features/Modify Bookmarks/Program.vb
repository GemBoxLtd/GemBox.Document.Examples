Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
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
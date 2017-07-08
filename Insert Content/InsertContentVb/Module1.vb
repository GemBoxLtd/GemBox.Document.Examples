Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        Dim bold = New CharacterFormat() With {.Bold = True}

        ' Create the whole document using streamlike API.
        document.Content.Start.
            LoadText("First paragraph.").
            InsertRange(New Paragraph(document, "Second paragraph.").Content).
            LoadText(vbLf & "Third bold paragraph.", bold)

        ' Prepend text to second paragraph.
        document.Sections(0).Blocks(1).Content.Start.LoadText("Some prefix (")

        ' Append text to second paragraph.
        document.Sections(0).Blocks(1).Content.End.LoadText(") some suffix.")

        ' Append text formatted using HTML tags.
        document.Sections(0).Blocks(2).Content.End.LoadText("<p>Fourth paragraph is added as <b>HTML</b> content.</p>", LoadOptions.HtmlDefault)

        document.Save("Insert Content.docx")

    End Sub

End Module
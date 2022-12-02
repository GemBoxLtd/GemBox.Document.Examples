Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        ' Create the whole document using fluent API.
        document.Content.Start _
            .LoadText("First paragraph.") _
            .InsertRange(New Paragraph(document, "Second paragraph.").Content) _
            .LoadText(vbLf) _
            .LoadText("Paragraph with bold text.", New CharacterFormat() With {.Bold = True})

        Dim section = document.Sections(0)

        ' Prepend text to second paragraph.
        section.Blocks(1).Content.Start.LoadText(" Some Prefix ", New CharacterFormat() With {.Subscript = True})

        ' Append text to second paragraph.
        section.Blocks(1).Content.End.LoadText(" Some Suffix ", New CharacterFormat() With {.Superscript = True})

        ' Insert HTML paragraph before third paragraph.
        section.Blocks(2).Content.Start.LoadText("<p style='font:italic 11pt Calibri;color:royalblue;'>Paragraph from HTML content with blue and italic text.</p>",
            New HtmlLoadOptions())

        ' Insert RTF paragraph after fourth paragraph.
        section.Blocks(3).Content.End.LoadText("{\rtf1\ansi\deff0{\colortbl ;\red255\green128\blue64;}\cf1 Paragraph from RTF content with orange text.\par\par}",
            New RtfLoadOptions())

        document.Save("Insert Content.docx")

    End Sub
End Module
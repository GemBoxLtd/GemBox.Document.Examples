Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        ' Set the content for the whole document.
        document.Content.LoadText("First paragraph." + vbLf + "Second paragraph." + vbLf + "Third paragraph." + vbLf + "Fourth paragraph." + vbLf + "Fifth paragraph.")

        Dim section = document.Sections(0)

        ' Set the content for 1st paragraph using plain text.
        section.Blocks(0).Content.LoadText("Paragraph with plain text.")

        ' Set the content for 2nd paragraph using specified formatting.
        section.Blocks(1).Content.LoadText("Paragraph with red and bold text.",
            New CharacterFormat() With {.FontColor = Color.Red, .Bold = True})

        Dim sourceRange As New ContentRange(
            section.Blocks(0).Content.Start,
            section.Blocks(1).Content.End)

        Dim destinationRange As New ContentRange(
            section.Blocks(2).Content.Start,
            section.Blocks(3).Content.End)

        ' Set the content for 3rd and 4th paragraph to be the same as the content of 1st and 2nd paragraph.
        destinationRange.Set(sourceRange)

        ' Set the content for 5th paragraph using HTML text.
        section.Blocks(4).Content.LoadText("<p style='font:11pt Calibri;color:blue;'>Paragraph with blue text that has <i>italic part</i>.</p>",
            New HtmlLoadOptions())

        document.Save("Set Content.docx")

    End Sub
End Module

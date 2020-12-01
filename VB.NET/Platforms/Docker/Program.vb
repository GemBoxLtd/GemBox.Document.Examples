Imports GemBox.Document

Module Program

    Sub Main()

        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Create new document.
        Dim document As New DocumentModel()

        ' Add sample text.
        document.Content.End _
            .LoadText("Lorem Ipsum" + vbLf, New CharacterFormat() With {.FontColor = Color.Red, .Bold = True}) _
            .LoadText("Lorem Ipsum" + vbLf, New CharacterFormat() With {.FontColor = Color.Green, .Italic = True}) _
            .LoadText("Lorem Ipsum" + vbLf, New CharacterFormat() With {.FontColor = Color.Blue, .UnderlineStyle = UnderlineType.Single})

        ' Add sample image.
        document.Content.End _
            .InsertRange(New Picture(document, "Dices.png").Content)

        ' Save document in DOCX and PDF format.
        document.Save("output.docx")
        document.Save("output.pdf")
    End Sub
End Module
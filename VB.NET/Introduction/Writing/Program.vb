Imports GemBox.Document

Module Program

    Sub Main()

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Create new empty document.
        Dim document As New DocumentModel()

        ' Add new section with two paragraphs, containing some text and symbols.
        document.Sections.Add(
            New Section(document,
                New Paragraph(document,
                    New Run(document, "This is our first paragraph with symbols added on a new line."),
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, ChrW(&HFC) & ChrW(&HF0) & ChrW(&H32)) With {.CharacterFormat = New CharacterFormat() With {.FontName = "Wingdings", .Size = 48}}),
                New Paragraph(document, "This is our second paragraph.")))

        ' Save Word document to file's path.
        document.Save("Writing.docx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Create new empty document.
        Dim document As New DocumentModel()

        ' Add plain text to document.
        document.Content.LoadText("This is a plain text.", New CharacterFormat() With {.FontName = "Arial"})

        ' Insert RTF formatted text at the beginning of the document.
        Dim content = document.Content.Start.LoadText("{\rtf1\ansi\deff0{\fonttbl{\f0 Arial Black;}}{\colortbl ;\red255\green128\blue64;}\f0\cf1 This is rich formatted text.}",
            LoadOptions.RtfDefault)

        ' Insert HTML formatted text after the previous text.
        content.LoadText("<p style='font-family:Arial Narrow;color:royalblue;'>This is another rich formatted text.</p>",
            LoadOptions.HtmlDefault)

        ' Save Word document to file's path.
        document.Save("WritingWithRichText.docx")
    End Sub

End Module

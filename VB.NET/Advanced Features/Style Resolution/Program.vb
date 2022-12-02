Imports System.Linq
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        ' Document's default font size is 8pt.
        document.DefaultCharacterFormat.Size = 8

        ' Style's font size is 24pt.
        Dim largeFont As New CharacterStyle("Large Font") With {.CharacterFormat = New CharacterFormat() With {.Size = 24}}
        document.Styles.Add(largeFont)

        Dim section As New Section(document)
        document.Sections.Add(section)

        Dim paragraph As New Paragraph(document,
            New Run(document, "Large text that has 'Large Font' style.") With
            {
                .CharacterFormat = New CharacterFormat() With {.Style = largeFont}
            },
            New SpecialCharacter(document, SpecialCharacterType.LineBreak),
            New Run(document, "Medium text that has both style and direct formatting; direct formatting has precedence over style's formatting.") With
            {
                .CharacterFormat = New CharacterFormat() With {.Style = largeFont, .Size = 12}
            },
            New SpecialCharacter(document, SpecialCharacterType.LineBreak),
            New Run(document, "Small text that uses document's default formatting."))

        section.Blocks.Add(paragraph)

        ' Write elements resolved font size values.
        For Each run As Run In document.GetChildElements(True, ElementType.Run).ToArray()
            section.Blocks.Add(New Paragraph(document, $"Font size: {run.CharacterFormat.Size} points. Text: {run.Text}"))
        Next

        document.Save("Style Resolution.docx")

    End Sub
End Module
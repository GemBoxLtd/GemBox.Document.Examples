Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        document.Sections.Add(
            New Section(document,
                New Paragraph(document,
                    New Run(document, "English: Hello"),
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Russian: "),
                    New Run(document, New String(New Char() {ChrW(&H417), ChrW(&H434), ChrW(&H440), ChrW(&H430), ChrW(&H432), ChrW(&H441), ChrW(&H442), ChrW(&H432), ChrW(&H443), ChrW(&H439), ChrW(&H442), ChrW(&H435)})),
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Chinese: "),
                    New Run(document, New String(New Char() {ChrW(&H4F60), ChrW(&H597D)}))),
                New Paragraph(document, "In order to see Russian and Chinese characters you need to have appropriate fonts on your machine.")))

        document.Save("Writing.docx")

    End Sub

End Module
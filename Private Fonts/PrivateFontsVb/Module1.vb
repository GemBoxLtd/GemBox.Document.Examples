Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        Dim pathToResources As String = "Resources"

        FontSettings.FontsBaseDirectory = pathToResources

        document.DefaultCharacterFormat = New CharacterFormat() With
        {
            .FontName = "Almonte Snow",
            .Size = 16
        }

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "Hello World!")))

        document.Save("Private Fonts.pdf")

    End Sub

End Module
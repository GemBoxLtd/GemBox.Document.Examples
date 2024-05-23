Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim section As New Section(document,
            New Paragraph(document,
                New Run(document, "First line."),
                New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                New Run(document, "Next line."),
                New SpecialCharacter(document, SpecialCharacterType.ColumnBreak),
                New Run(document, "Next column."),
                New SpecialCharacter(document, SpecialCharacterType.PageBreak),
                New Run(document, "Next page.")))

        section.PageSetup.TextColumns = New TextColumnCollection(2)

        document.Sections.Add(section)

        document.Save("Breaks.docx")

    End Sub
End Module

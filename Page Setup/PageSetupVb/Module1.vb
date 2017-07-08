Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        Dim section As New Section(document,
            New Paragraph(document,
                New Run(document, "First line"),
                New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                New Run(document, "Second line"),
                New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                New Run(document, "Third line")),
            New Paragraph(document,
                New SpecialCharacter(document, SpecialCharacterType.ColumnBreak),
                New Run(document, "First line"),
                New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                New Run(document, "Second line"),
                New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                New Run(document, "Third line")))

        Dim pageSetup As PageSetup = section.PageSetup

        ' Specify text columns.
        pageSetup.TextColumns = New TextColumnCollection(2) With {
            .LineBetween = True,
            .EvenlySpaced = False
        }

        pageSetup.TextColumns(0).Width = LengthUnitConverter.Convert(1, LengthUnit.Inch, LengthUnit.Point)
        pageSetup.TextColumns(1).Width = LengthUnitConverter.Convert(2.3, LengthUnit.Inch, LengthUnit.Point)

        ' Specify paper type.
        pageSetup.PaperType = PaperType.A5

        document.Sections.Add(section)

        ' Specify line numbering.
        document.Sections.Add(
            New Section(document,
                New Paragraph(document,
                    New Run(document, "First line"),
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Second line"),
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Third line"))) With {
                        .PageSetup = New PageSetup() With {
                            .PaperType = PaperType.A5,
                            .LineNumberRestartSetting = LineNumberRestartSetting.NewPage
         }})

        document.Save("Page Setup.docx")

    End Sub

End Module
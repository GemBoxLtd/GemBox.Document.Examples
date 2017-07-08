Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "Text is centered") With {
                    .ParagraphFormat = New ParagraphFormat() With {
                        .Alignment = HorizontalAlignment.Center
                    }
                },
                New Paragraph(document, "This paragraph has the following properties: Left indentation is 10 points, right indentation is 10 points, hanging indentation is 20 points, line spacing is exactly 20 points, space before and space after are 20 points.") With {
                    .ParagraphFormat = New ParagraphFormat() With {
                        .LeftIndentation = 10,
                        .RightIndentation = 10,
                        .SpecialIndentation = 20,
                        .LineSpacing = 20,
                        .LineSpacingRule = LineSpacingRule.Exactly,
                        .SpaceBefore = 20,
                        .SpaceAfter = 20
                    }
                },
                New Paragraph(document, "This paragraph has the following properties: Left indentation is 25 points, right indentation is 25 points, first line indentation is 25 points and line spacing is at least 10 points.") With {
                    .ParagraphFormat = New ParagraphFormat() With {
                        .LeftIndentation = 25,
                        .RightIndentation = 25,
                        .SpecialIndentation = -25,
                        .LineSpacing = 10,
                        .LineSpacingRule = LineSpacingRule.AtLeast
                }}))

        Dim paragraphWithBorders As New Paragraph(document, "The following paragraph is surrounded with the borders.")
        paragraphWithBorders.ParagraphFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 2)
        document.Sections(0).Blocks.Add(paragraphWithBorders)

        document.Save("Paragraph Formatting.docx")

    End Sub

End Module
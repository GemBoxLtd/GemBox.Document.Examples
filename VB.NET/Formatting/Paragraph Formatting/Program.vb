Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "This paragraph has centered text.") With
                {
                    .ParagraphFormat = New ParagraphFormat() With
                    {
                        .Alignment = HorizontalAlignment.Center
                    }
                },
                New Paragraph(document, "This paragraph has the following properties:" & vbLf & "Left indentation is 10 points." & vbLf & "Right indentation is 20 points." & vbLf & "Hanging indentation is 30 points.") With
                {
                    .ParagraphFormat = New ParagraphFormat() With
                    {
                        .LeftIndentation = 10,
                        .RightIndentation = 20,
                        .SpecialIndentation = 30
                    }
                },
                New Paragraph(document, "This paragraph has the following properties:" & vbLf & "First line indentation is 40 points." & vbLf & "Line spacing is exactly 30 points." & vbLf & "Space after and before are 10 points.") With
                {
                    .ParagraphFormat = New ParagraphFormat() With
                    {
                        .SpecialIndentation = -40,
                        .LineSpacing = 30,
                        .LineSpacingRule = LineSpacingRule.Exactly,
                        .SpaceAfter = 10,
                        .SpaceBefore = 10
                    }
                }))

        Dim paragraph As New Paragraph(document, "This paragraph has borders and background color.")
        paragraph.ParagraphFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, New Color(237, 125, 49), 2)
        paragraph.ParagraphFormat.BackgroundColor = New Color(251, 228, 213)
        document.Sections(0).Blocks.Add(paragraph)

        document.Save("Paragraph Formatting.docx")

    End Sub
End Module
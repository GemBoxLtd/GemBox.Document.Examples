Imports System.Globalization
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        document.DefaultCharacterFormat.FontName = "Arial"
        document.DefaultCharacterFormat.Size = 16

        Dim lineBreakElement As New SpecialCharacter(document, SpecialCharacterType.LineBreak)

        document.Sections.Add(
            New Section(document,
                New Paragraph(document,
                    New Run(document, "All caps") With {.CharacterFormat = New CharacterFormat With {.AllCaps = True}},
                    lineBreakElement,
                    New Run(document, "Text with background color") With {.CharacterFormat = New CharacterFormat With {.BackgroundColor = Color.Cyan}},
                    lineBreakElement.Clone(),
                    New Run(document, "Bold text") With {.CharacterFormat = New CharacterFormat With {.Bold = True}},
                    lineBreakElement.Clone(),
                    New Run(document, "Text with borders") With {.CharacterFormat = New CharacterFormat With {.Border = New SingleBorder(BorderStyle.Single, Color.Red, 1)}},
                    lineBreakElement.Clone(),
                    New Run(document, "Double strikethrough text") With {.CharacterFormat = New CharacterFormat With {.DoubleStrikethrough = True}},
                    lineBreakElement.Clone(),
                    New Run(document, "Blue text") With {.CharacterFormat = New CharacterFormat With {.FontColor = Color.Blue}},
                    lineBreakElement.Clone(),
                    New Run(document, "Text with 'Consolas' font") With {.CharacterFormat = New CharacterFormat With {.FontName = "Consolas"}},
                    lineBreakElement.Clone(),
                    New Run(document, "Hidden text") With {.CharacterFormat = New CharacterFormat With {.Hidden = True}},
                    lineBreakElement.Clone(),
                    New Run(document, "Text with highlight color") With {.CharacterFormat = New CharacterFormat With {.HighlightColor = Color.Yellow}},
                    lineBreakElement.Clone(),
                    New Run(document, "Italic text") With {.CharacterFormat = New CharacterFormat With {.Italic = True}},
                    lineBreakElement.Clone(),
                    New Run(document, "Kerning is 15 points") With {.CharacterFormat = New CharacterFormat With {.Kerning = 15}},
                    lineBreakElement.Clone(),
                    New Run(document, "Position is 3 points") With {.CharacterFormat = New CharacterFormat With {.Position = 3}},
                    lineBreakElement.Clone(),
                    New Run(document, "Scale is 125%") With {.CharacterFormat = New CharacterFormat With {.Scaling = 125}},
                    lineBreakElement.Clone(),
                    New Run(document, "Font size is 24 points") With {.CharacterFormat = New CharacterFormat With {.Size = 24}},
                    lineBreakElement.Clone(),
                    New Run(document, "Small caps") With {.CharacterFormat = New CharacterFormat With {.SmallCaps = True}},
                    lineBreakElement.Clone(),
                    New Run(document, "Spacing is 3 point") With {.CharacterFormat = New CharacterFormat With {.Spacing = 3}},
                    lineBreakElement.Clone(),
                    New Run(document, "Strikethrough text") With {.CharacterFormat = New CharacterFormat With {.Strikethrough = True}},
                    lineBreakElement.Clone(),
                    New Run(document, "Subscript text") With {.CharacterFormat = New CharacterFormat With {.Subscript = True}},
                    lineBreakElement.Clone(),
                    New Run(document, "Superscript text") With {.CharacterFormat = New CharacterFormat With {.Superscript = True}},
                    lineBreakElement.Clone(),
                    New Run(document, "Underline color is orange") With {.CharacterFormat = New CharacterFormat With {.UnderlineColor = Color.Orange, .UnderlineStyle = UnderlineType.Single}},
                    lineBreakElement.Clone(),
                    New Run(document, "Underline style is double") With {.CharacterFormat = New CharacterFormat With {.UnderlineStyle = UnderlineType.Double}},
                    lineBreakElement.Clone(),
                    New Field(document, FieldType.Date, "\@ ""dddd, d. MMMM yyyy""") With {.CharacterFormat = New CharacterFormat With {.Language = CultureInfo.GetCultureInfo("de-DE")}})))

        document.Save("Character Formatting.docx")

    End Sub
End Module
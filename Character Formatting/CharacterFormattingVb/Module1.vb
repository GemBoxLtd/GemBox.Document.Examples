Imports System
Imports System.Globalization
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        document.Sections.Add(
            New Section(document,
                New Paragraph(document,
                    New Run(document, "All caps: "),
                    New Run(document, "Capital letters") With {
                        .CharacterFormat = New CharacterFormat() With {
                            .AllCaps = True
                        }
                    },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Bold: "),
                    New Run(document, "Bold text") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Bold = True
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Border: "),
                    New Run(document, "Some text") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Border = New SingleBorder(BorderStyle.Single, Color.Black, 1)
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Double strikethrough: "),
                    New Run(document, "Some text") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .DoubleStrikethrough = True
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Font color: "),
                    New Run(document, "Blue text") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .FontColor = Color.Blue
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Font name: "),
                    New Run(document, "Arial Black") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .FontName = "Arial Black"
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Hidden: "),
                    New Run(document, "Hidden text") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Hidden = True
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Text highlight color: "),
                    New Run(document, "Yellow background color") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .HighlightColor = Color.Yellow
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Italic: "),
                    New Run(document, "Italic text") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Italic = True
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Kerning: "),
                    New Run(document, "Kerning is 15 points") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Kerning = 15
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Position: "),
                    New Run(document, "Position is 3 points") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Position = 3
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Scale: "),
                    New Run(document, "Scale is 125%") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Scaling = 125
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Font size: "),
                    New Run(document, "Font size is 14 points") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Size = 14
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Small caps: "),
                    New Run(document, "Some text") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .SmallCaps = True
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Spacing: "),
                    New Run(document, "Spacing is 1 point") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Spacing = 1
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Strikethrough: "),
                    New Run(document, "Some text") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Strikethrough = True
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Subscript: "),
                    New Run(document, "Some text") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Subscript = True
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Superscript: "),
                    New Run(document, "Some text") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Superscript = True
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Underline color: "),
                    New Run(document, "Underline color is blue") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .UnderlineColor = Color.Blue,
                                .UnderlineStyle = UnderlineType.Single
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Underline style: "),
                    New Run(document, "Underline style is double") With {
                            .CharacterFormat = New CharacterFormat() With {
                                .UnderlineStyle = UnderlineType.Double
                            }
                        },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Language: "),
                    New Field(document, FieldType.Date) With {
                            .CharacterFormat = New CharacterFormat() With {
                                .Language = CultureInfo.GetCultureInfo("de-DE")
                            }
                        })))

        document.Save("Character Formatting.docx")

    End Sub

End Module

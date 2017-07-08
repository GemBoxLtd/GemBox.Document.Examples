using System;
using System.Globalization;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        document.Sections.Add(
           new Section(document,
               new Paragraph(document,
                   new Run(document, "All caps: "),
                   new Run(document, "Capital letters")
                   {
                       CharacterFormat = new CharacterFormat()
                       {
                           AllCaps = true,
                       }
                   },
                   new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                   new Run(document, "Bold: "),
                   new Run(document, "Bold text")
                   {
                       CharacterFormat = new CharacterFormat()
                       {
                           Bold = true
                       }
                   },
                   new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                   new Run(document, "Border: "),
                   new Run(document, "Some text")
                   {
                       CharacterFormat = new CharacterFormat()
                       {
                           Border = new SingleBorder(BorderStyle.Single, Color.Black, 1)
                       }
                   },
                   new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                   new Run(document, "Double strikethrough: "),
                   new Run(document, "Some text")
                   {
                       CharacterFormat = new CharacterFormat()
                       {
                           DoubleStrikethrough = true
                       }
                   },
                   new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                   new Run(document, "Font color: "),
                   new Run(document, "Blue text")
                   {
                       CharacterFormat = new CharacterFormat()
                       {
                           FontColor = Color.Blue
                       }
                   },
                   new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                   new Run(document, "Font name: "),
                   new Run(document, "Arial Black")
                   {
                       CharacterFormat = new CharacterFormat()
                       {
                           FontName = "Arial Black"
                       }
                   },
                   new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                   new Run(document, "Hidden: "),
                   new Run(document, "Hidden text")
                   {
                       CharacterFormat = new CharacterFormat()
                       {
                           Hidden = true
                       }
                   },
                   new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                   new Run(document, "Text highlight color: "),
                   new Run(document, "Yellow background color")
                   {
                       CharacterFormat = new CharacterFormat()
                       {
                           HighlightColor = Color.Yellow
                       }
                   },
                   new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                   new Run(document, "Italic: "),
                   new Run(document, "Italic text")
                   {
                       CharacterFormat = new CharacterFormat()
                       {
                           Italic = true
                       }
                   },
                   new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                   new Run(document, "Kerning: "),
                    new Run(document, "Kerning is 15 points")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            Kerning = 15
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Position: "),
                    new Run(document, "Position is 3 points")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            Position = 3
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Scale: "),
                    new Run(document, "Scale is 125%")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            Scaling = 125
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Font size: "),
                    new Run(document, "Font size is 14 points")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            Size = 14
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Small caps: "),
                    new Run(document, "Some text")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            SmallCaps = true
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Spacing: "),
                    new Run(document, "Spacing is 1 point")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            Spacing = 1
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Strikethrough: "),
                    new Run(document, "Some text")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            Strikethrough = true
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Subscript: "),
                    new Run(document, "Some text")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            Subscript = true
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Superscript: "),
                    new Run(document, "Some text")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            Superscript = true
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Underline color: "),
                    new Run(document, "Underline color is blue")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            UnderlineColor = Color.Blue,
                            UnderlineStyle = UnderlineType.Single
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Underline style: "),
                    new Run(document, "Underline style is double")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            UnderlineStyle = UnderlineType.Double
                        }
                    },
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Language: "),
                    new Field(document, FieldType.Date)
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            Language = CultureInfo.GetCultureInfo("de-DE")
                        }
                    })));

        document.Save("Character Formatting.docx");
    }
}

using System.Globalization;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        document.DefaultCharacterFormat.FontName = "Arial";
        document.DefaultCharacterFormat.Size = 16;

        SpecialCharacter lineBreakElement = new SpecialCharacter(document, SpecialCharacterType.LineBreak);

        document.Sections.Add(
            new Section(document,
                new Paragraph(document,
                    new Run(document, "All caps") { CharacterFormat = { AllCaps = true } },
                    lineBreakElement,
                    new Run(document, "Text with background color") { CharacterFormat = { BackgroundColor = Color.Cyan } },
                    lineBreakElement.Clone(),
                    new Run(document, "Bold text") { CharacterFormat = { Bold = true } },
                    lineBreakElement.Clone(),
                    new Run(document, "Text with borders") { CharacterFormat = { Border = new SingleBorder(BorderStyle.Single, Color.Red, 1) } },
                    lineBreakElement.Clone(),
                    new Run(document, "Double strikethrough text") { CharacterFormat = { DoubleStrikethrough = true } },
                    lineBreakElement.Clone(),
                    new Run(document, "Blue text") { CharacterFormat = { FontColor = Color.Blue } },
                    lineBreakElement.Clone(),
                    new Run(document, "Text with 'Consolas' font") { CharacterFormat = { FontName = "Consolas" } },
                    lineBreakElement.Clone(),
                    new Run(document, "Hidden text") { CharacterFormat = { Hidden = true } },
                    lineBreakElement.Clone(),
                    new Run(document, "Text with highlight color") { CharacterFormat = { HighlightColor = Color.Yellow } },
                    lineBreakElement.Clone(),
                    new Run(document, "Italic text") { CharacterFormat = { Italic = true } },
                    lineBreakElement.Clone(),
                    new Run(document, "Kerning is 15 points") { CharacterFormat = { Kerning = 15 } },
                    lineBreakElement.Clone(),
                    new Run(document, "Position is 3 points") { CharacterFormat = { Position = 3 } },
                    lineBreakElement.Clone(),
                    new Run(document, "Scale is 125%") { CharacterFormat = { Scaling = 125 } },
                    lineBreakElement.Clone(),
                    new Run(document, "Font size is 24 points") { CharacterFormat = { Size = 24 } },
                    lineBreakElement.Clone(),
                    new Run(document, "Small caps") { CharacterFormat = { SmallCaps = true } },
                    lineBreakElement.Clone(),
                    new Run(document, "Spacing is 3 point") { CharacterFormat = { Spacing = 3 } },
                    lineBreakElement.Clone(),
                    new Run(document, "Strikethrough text") { CharacterFormat = { Strikethrough = true } },
                    lineBreakElement.Clone(),
                    new Run(document, "Subscript text") { CharacterFormat = { Subscript = true } },
                    lineBreakElement.Clone(),
                    new Run(document, "Superscript text") { CharacterFormat = { Superscript = true } },
                    lineBreakElement.Clone(),
                    new Run(document, "Underline color is orange") { CharacterFormat = { UnderlineColor = Color.Orange, UnderlineStyle = UnderlineType.Single } },
                    lineBreakElement.Clone(),
                    new Run(document, "Underline style is double") { CharacterFormat = { UnderlineStyle = UnderlineType.Double } },
                    lineBreakElement.Clone(),
                    new Field(document, FieldType.Date, @"\@ ""dddd, d. MMMM yyyy""") { CharacterFormat = { Language = CultureInfo.GetCultureInfo("de-DE") } })));

        document.Save("Character Formatting.docx");
    }
}
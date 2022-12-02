using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "This paragraph has centered text.")
                {
                    ParagraphFormat =
                    {
                        Alignment = HorizontalAlignment.Center
                    }
                },
                new Paragraph(document, "This paragraph has the following properties:\nLeft indentation is 10 points.\nRight indentation is 20 points.\nHanging indentation is 30 points.")
                {
                    ParagraphFormat =
                    {
                        LeftIndentation = 10,
                        RightIndentation = 20,
                        SpecialIndentation = 30
                    }
                },
                new Paragraph(document, "This paragraph has the following properties:\nFirst line indentation is 40 points.\nLine spacing is exactly 30 points.\nSpace after and before are 10 points.")
                {
                    ParagraphFormat = new ParagraphFormat
                    {
                        SpecialIndentation = -40,
                        LineSpacing = 30,
                        LineSpacingRule = LineSpacingRule.Exactly,
                        SpaceBefore = 10,
                        SpaceAfter = 10
                    }
                }));

        Paragraph paragraph = new Paragraph(document, "This paragraph has borders and background color.");
        paragraph.ParagraphFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, new Color(237, 125, 49), 2);
        paragraph.ParagraphFormat.BackgroundColor = new Color(251, 228, 213);
        document.Sections[0].Blocks.Add(paragraph);

        document.Save("Paragraph Formatting.docx");
    }
}
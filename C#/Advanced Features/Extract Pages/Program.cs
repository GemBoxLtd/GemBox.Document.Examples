using GemBox.Document;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Input.docx");

        var paginator = document.GetPaginator();

        var secondPage = paginator.Pages[1];

        secondPage.Save("SecondPage.docx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Input.docx");

        var paginator = document.GetPaginator();

        for (var i = 0; i < paginator.Pages.Count; i++)
        {
            var page = paginator.Pages[i];
            var pageRange = page.Range;
            var start = pageRange.Start[0];

            var textBox = CreateTextBox(document, i);
            start.InsertRange(textBox.Content);
        }

        document.Save("Output.docx");
    }

    // A floating textbox that will be inserted at the start of every page.
    private static TextBox CreateTextBox(DocumentModel document, int page)
    {
        var run = new Run(document, "Inserted textbox on page " + (page + 1));
        run.CharacterFormat.Size = 25;
        run.CharacterFormat.FontColor = Color.White;

        var textBox = new TextBox(document, new FloatingLayout(
            new HorizontalPosition(-340, LengthUnit.Point, HorizontalPositionAnchor.RightMargin),
            new VerticalPosition(0, LengthUnit.Point, VerticalPositionAnchor.Margin),
            new Size(340, 45, LengthUnit.Point))
        { WrappingStyle = TextWrappingStyle.InFrontOfText });
        textBox.Fill.SetSolid(new Color(0x4472C4));
        textBox.Blocks.Add(new Paragraph(document, run));

        return textBox;
    }
}

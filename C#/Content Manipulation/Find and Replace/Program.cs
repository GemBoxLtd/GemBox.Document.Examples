using System;
using System.Linq;
using System.Text.RegularExpressions;
using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        var document = DocumentModel.Load("FindAndReplaceText.docx");

        // The easiest way how you can find and replace text is with "Replace" method.
        document.Content.Replace("%FirstName%", "John");
        document.Content.Replace("%LastName%", "Doe");

        // Another way would be to use Regex.
        document.Content.Replace(new Regex("%DATE%", RegexOptions.IgnoreCase),
            DateTime.Today.ToLongDateString());

        document.Content.Replace(new Regex("%.*?%"), range =>
        {
            string value = null;
            switch (range.ToString())
            {
                case "%Address%": value = "240 Old Country Road"; break;
                case "%City%": value = "Springfield"; break;
                case "%State%": value = "IL"; break;
                case "%Country%": value = "USA"; break;
            };

            if (string.IsNullOrEmpty(value))
                return range;

            var format = ((Run)range.Start.Parent).CharacterFormat;
            var run = new Run(document, value) { CharacterFormat = format.Clone() };
            return run.Content;
        });

        // You can also search for placeholder text with the "Find" method and then achieve a
        // more complex replacement, like the following which has a replace text with different formatting.
        // Notice that the "Reverse" extension method is used here to avoid a possible invalid state because
        // the replacements are done while iterating through the document's content.
        foreach (ContentRange searchedContent in document.Content.Find("%Price%").Reverse())
        {
            ContentRange replacedContent = searchedContent.LoadText("$",
                new CharacterFormat() { Size = 14, FontColor = Color.Blue, Bold = true });
            replacedContent.End.LoadText("100.00",
                new CharacterFormat() { Size = 11, FontColor = Color.Purple, Italic = true });
        }

        // Another more complex replacement in which searched text is replaced with a hyperlink.
        foreach (ContentRange searchedContent in document.Content.Find("%Email%").Reverse())
        {
            Hyperlink emailLink = new Hyperlink(document, "mailto:john.doe@example.com", "John.Doe@example.com");
            searchedContent.Set(emailLink.Content);
        }

        // You can also find and highlight text by specifying "HighlightColor" of replacement text.
        foreach (ContentRange searchedContent in document.Content.Find("membership").Reverse())
        {
            var highlightedText = new Run(document, "membership");
            highlightedText.CharacterFormat = ((Run)searchedContent.Start.Parent).CharacterFormat.Clone();
            highlightedText.CharacterFormat.HighlightColor = Color.Yellow;
            searchedContent.Set(highlightedText.Content);
        }

        document.Save("FoundAndReplacedText.docx");
    }

    static void Example2()
    {
        var document = DocumentModel.Load("FindAndReplaceContent.docx");

        var dummyText = "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.";

        // Find an image placeholder.
        var picturePlaceholder = document.Content.Find("%Portrait%").First();
        var picture = new Picture(document, "avatar.png");

        // Replace the placeholder text with the image.
        picturePlaceholder.Set(picture.Content);

        // Find an HTML placeholder.
        var htmlPlaceholder = document.Content.Find("%AboutMe%").First();
        var html =
$@"<ul style='font:11pt Calibri;'>
    <li style='color:red;'>{dummyText}</li>
    <li style='color:green;'>{dummyText}</li>
    <li style='color:blue;'>{dummyText}</li>
</ul>";

        // Replace the placeholder text with HTML formatted text.
        htmlPlaceholder.LoadText(html, new HtmlLoadOptions());

        // Find a table placeholder.
        var tablePlaceholder = document.Content.Find("%JobHistory%").First();

        var table = new Table(document,
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "2021 - 2030")),
                new TableCell(document, new Paragraph(document, dummyText))),
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "2011 - 2020")),
                new TableCell(document, new Paragraph(document, dummyText))),
            new TableRow(document,
                new TableCell(document, new Paragraph(document, "2001 - 2010")),
                new TableCell(document, new Paragraph(document, dummyText))));

        table.Columns.Add(new TableColumn(70));
        table.Columns.Add(new TableColumn(250));
        table.TableFormat.AutomaticallyResizeToFitContents = false;

        // Delete the placeholder text and insert the table before it.
        tablePlaceholder = tablePlaceholder.LoadText(string.Empty);
        tablePlaceholder.Start.InsertRange(table.Content);

        document.Save("FoundAndReplacedContent.docx");
    }
}
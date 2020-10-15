using System;
using System.Linq;
using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
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
        document.Content.Replace("%Date%", DateTime.Today.ToLongDateString());

        // You can also find and highlight text by specifying "HighlightColor" of replacement text.
        document.Content.Replace("membership", "membership", new CharacterFormat() { HighlightColor = Color.Yellow });

        // You can also search for placeholder text with "Find" method and then achieve more complex replacement.
        // The "Reverse" extension method is used when iterating through document's content
        // to avoid any possible invalid state when replacing the content within iteration.

        // Replace text with text that has different formatting.
        foreach (ContentRange searchedContent in document.Content.Find("%Price%").Reverse())
        {
            ContentRange replacedContent = searchedContent.LoadText("$",
                new CharacterFormat() { Size = 14, FontColor = Color.Purple });
            replacedContent.End.LoadText("100.00",
                new CharacterFormat() { Size = 11, FontColor = Color.Purple });
        }

        // Replace text with hyperlink.
        foreach (ContentRange searchedContent in document.Content.Find("%Email%").Reverse())
        {
            Hyperlink emailLink = new Hyperlink(document, "mailto:john.doe@example.com", "John.Doe@example.com");
            searchedContent.Set(emailLink.Content);
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
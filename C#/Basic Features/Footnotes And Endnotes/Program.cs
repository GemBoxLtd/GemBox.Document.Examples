using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        // Set footnotes and endnotes number style for all sections in the document.
        document.Settings.Footnote.NumberStyle = NumberStyle.LowerLetter;
        document.Settings.Endnote.NumberStyle = NumberStyle.LowerRoman;

        var section = new Section(document);
        document.Sections.Add(section);

        // Set footnotes number style for the current section.
        section.FootnoteSettings.NumberStyle = NumberStyle.Decimal;

        section.Blocks.Add(
            new Paragraph(document,
                new Run(document, "GemBox.Document"),
                new Note(document, NoteType.Footnote,
                    new Paragraph(document,
                        new Run(document, "Read more about GemBox.Document on "),
                        new Hyperlink(document, "https://www.gemboxsoftware.com/document", "overview page"),
                        new Run(document, "."))),
                new Run(document, " is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT)"),
                new Note(document, NoteType.Footnote, "Image formats like PNG, JPEG, GIF, BMP, TIFF and WMP are also supported."),
                new Run(document, " from .NET"),
                new Note(document, NoteType.Endnote, "GemBox.Document works on .NET Framework 3.5 or higher and platforms that implement .NET Standard 2.0 or higher."),
                new Run(document, " applications in a simple and efficient way.")));

        section = new Section(document);
        document.Sections.Add(section);

        // Set endnotes number style for the current section.
        section.EndnoteSettings.NumberStyle = NumberStyle.UpperRoman;

        section.Blocks.Add(
            new Paragraph(document,
                new Run(document, "The latest version of GemBox.Document can be downloaded from "),
                new Hyperlink(document, "https://www.gemboxsoftware.com/document/free-version", "here"),
                new Note(document, NoteType.Endnote,
                    new Paragraph(document,
                        new Run(document, "The latest fixes for all GemBox.Document versions can be downloaded from "),
                        new Hyperlink(document, "https://www.gemboxsoftware.com/document/downloads/BugFixes.htm", "here"),
                        new Run(document, "."))),
                new Run(document, ".")));

        document.Save("Footnotes and Endnotes.docx");
    }
}
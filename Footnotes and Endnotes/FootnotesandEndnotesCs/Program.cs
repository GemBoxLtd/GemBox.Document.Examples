using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        // All sections added to the document will have footnotes
        // numbered using lower letters.
        document.Settings.Footnote.NumberStyle = NumberStyle.LowerLetter;

        // All sections added to the document will have endnotes
        // numbered using lower Roman.
        document.Settings.Endnote.NumberStyle = NumberStyle.LowerRoman;

        Section section = new Section(document);
        document.Sections.Add(section);

        // All footnotes in this section will be numbered using digits.
        // All endnotes in this section will be numbered using lower Roman 
        // (taken from document settings).
        section.FootnoteSettings.NumberStyle = NumberStyle.Decimal;

        section.Blocks.Add(new Paragraph(document,
            new Run(document, "GemBox.Document"),
            new Note(document, NoteType.Footnote,
                new Paragraph(document,
                    new Run(document, "More info at "),
                    new Hyperlink(document,
                        "https://www.gemboxsoftware.com/document/overview",
                        "our website"),
                    new Run(document, "."))),
            new Run(document, " is a .NET component that enables developers to read,"),
            new Run(document, " write, convert and print document files"),
            new Run(document, " (DOCX, DOC, PDF, HTML, XPS, RTF and TXT)"),
            new Note(document, NoteType.Endnote,
                "Image formats like PNG, JPEG, BMP and TIFF are also supported."),
            new Run(document, " from .NET"),
            new Note(document, NoteType.Footnote,
                "Minimum required .NET Framework version is 3.0."),
            new Run(document, " applications in a simple and efficient way.")));

        section = new Section(document);
        document.Sections.Add(section);

        // All endnotes in this section will be numbered using upper Roman.
        // All footnotes in this section will be numbered using lower letters
        // (taken from document settings).
        section.EndnoteSettings.NumberStyle = NumberStyle.UpperRoman;

        section.Blocks.Add(new Paragraph(document,
            new Run(document, "All versions of GemBox.Document come in "),
            new Run(document, "the Microsoft Windows Installer"),
            new Note(document, NoteType.Footnote,
                ".MSI file format."),
            new Run(document, " format, which ensures you can easily install /"),
            new Run(document, " uninstall and use multiple GemBox.Document versions"),
            new Note(document, NoteType.Endnote,
                new Paragraph(document,
                    new Run(document, "You can get the latest version "),
                    new Hyperlink(document,
                        "https://www.gemboxsoftware.com/document/free-version",
                        "here"),
                    new Run(document, "."))),
            new Run(document, " on the same machine.")));

        document.Sections.Add(
            new Section(document,
                new Paragraph(document,
                    new Run(document, "GemBox.Document Free"),
                    new Note(document, NoteType.Endnote,
                        "Limited to 20 paragraphs."),
                    new Run(document, " is free of charge while GemBox.Document"),
                    new Run(document, " Professional is a commercial version"),
                    new Run(document, " licensed per developer."),
                    new Run(document, " Server deployment is royalty free."))));

        document.Save("Footnotes and Endnotes.docx");
    }
}

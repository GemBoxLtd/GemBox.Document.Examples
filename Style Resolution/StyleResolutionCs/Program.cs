using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        // Default font size is 12
        document.DefaultCharacterFormat.Size = 12;

        CharacterStyle largeSize = new CharacterStyle("Large Font Size")
        {
            CharacterFormat = new CharacterFormat()
            {
                Size = 24
            }
        };

        // Runs with the following style will have font size 24
        document.Styles.Add(largeSize);

        // Fill section with test data.
        document.Sections.Add(
            new Section(document,
                new Paragraph(document,
                    new Run(document, "GemBox.Document component")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            // Style
                            Style = largeSize
                        }
                    }),
                new Paragraph(document,
                    new Run(document, "GemBox.")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            // Style and direct formatting (direct formatting has precedence over style)
                            Style = largeSize,
                            Size = 16
                        }
                    },
                    new Run(document, "Document")
                    {
                        CharacterFormat = new CharacterFormat()
                        {
                            // Direct formatting
                            Size = 16
                        }
                    },// Default style                
                    new Run(document, " is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way."))));

        // Fill section with results.
        Section section = new Section(document);
        Paragraph para = new Paragraph(document);
        section.Blocks.Add(para);

        foreach (Run run in document.GetChildElements(true, ElementType.Run))
        {
            para.Inlines.Add(
                new Run(document, "Font size: " + run.CharacterFormat.Size + " points, Text: '" + run.Text + "'"));

            para.Inlines.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));
        }

        document.Sections.Add(section);

        document.Save("Style Resolution.docx");
    }
}

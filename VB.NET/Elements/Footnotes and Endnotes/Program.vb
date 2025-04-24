Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        ' Set footnotes and endnotes number style for all sections in the document.
        document.Settings.Footnote.NumberStyle = NumberStyle.LowerLetter
        document.Settings.Endnote.NumberStyle = NumberStyle.LowerRoman

        Dim section As New Section(document)
        document.Sections.Add(section)

        ' Set footnotes number style for the current section.
        section.FootnoteSettings.NumberStyle = NumberStyle.Decimal

        section.Blocks.Add(
            New Paragraph(document,
                New Run(document, "GemBox.Document"),
                New Note(document, NoteType.Footnote,
                    New Paragraph(document,
                        New Run(document, "Read more about GemBox.Document on "),
                        New Hyperlink(document, "https://www.gemboxsoftware.com/document", "overview page"),
                        New Run(document, "."))),
                New Run(document, " is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF, ODT, and TXT)"),
                New Note(document, NoteType.Footnote, "Image formats like PNG, JPEG, GIF, BMP, TIFF, and WMP are also supported."),
                New Run(document, " from .NET"),
                New Note(document, NoteType.Endnote, "GemBox.Document works on platforms that implement .NET Standard 2.0 or higher."),
                New Run(document, " applications in a simple and efficient way.")))

        section = New Section(document)
        document.Sections.Add(section)

        ' Set endnotes number style for the current section.
        section.EndnoteSettings.NumberStyle = NumberStyle.UpperRoman

        section.Blocks.Add(
            New Paragraph(document,
                New Run(document, "The latest version of GemBox.Document can be downloaded from "),
                New Hyperlink(document, "https://www.gemboxsoftware.com/document/free-version", "here"),
                New Note(document, NoteType.Endnote,
                    New Paragraph(document,
                        New Run(document, "The latest fixes for all GemBox.Document versions can be downloaded from "),
                        New Hyperlink(document, "https://www.gemboxsoftware.com/document/downloads/", "here"),
                        New Run(document, "."))),
                New Run(document, ".")))

        document.Save("Footnotes and Endnotes.docx")

    End Sub
End Module

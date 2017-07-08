Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        ' All sections added to the document will have footnotes
        ' numbered using lower letters.
        document.Settings.Footnote.NumberStyle = NumberStyle.LowerLetter

        ' All sections added to the document will have endnotes
        ' numbered using lower Roman.
        document.Settings.Endnote.NumberStyle = NumberStyle.LowerRoman

        Dim section As New Section(document)
        document.Sections.Add(section)

        ' All footnotes in this section will be numbered using digits.
        ' All endnotes in this section will be numbered using lower Roman 
        ' (taken from document settings).
        section.FootnoteSettings.NumberStyle = NumberStyle.Decimal

        section.Blocks.Add(New Paragraph(document,
            New Run(document, "GemBox.Document"),
            New Note(document, NoteType.Footnote,
                New Paragraph(document,
                    New Run(document, "More info at "),
                    New Hyperlink(document,
                        "https://www.gemboxsoftware.com/document/overview",
                        "our website"),
                    New Run(document, "."))),
            New Run(document, " is a .NET component that enables developers to read,"),
            New Run(document, " write, convert and print document files"),
            New Run(document, " (DOCX, DOC, PDF, HTML, XPS, RTF and TXT)"),
            New Note(document, NoteType.Endnote,
                "Image formats like PNG, JPEG, BMP and TIFF are also supported."),
            New Run(document, " from .NET"),
            New Note(document, NoteType.Footnote,
                "Minimum required .NET Framework version is 3.0."),
            New Run(document, " applications in a simple and efficient way.")))

        section = New Section(document)
        document.Sections.Add(section)

        ' All endnotes in this section will be numbered using upper Roman.
        ' All footnotes in this section will be numbered using lower letters
        ' (taken from document settings).
        section.EndnoteSettings.NumberStyle = NumberStyle.UpperRoman

        section.Blocks.Add(New Paragraph(document,
            New Run(document, "All versions of GemBox.Document come in "),
            New Run(document, "the Microsoft Windows Installer"),
            New Note(document, NoteType.Footnote,
                ".MSI file format."),
            New Run(document, " format, which ensures you can easily install /"),
            New Run(document, " uninstall and use multiple GemBox.Document versions"),
            New Note(document, NoteType.Endnote,
                New Paragraph(document,
                    New Run(document, "You can get the latest version "),
                    New Hyperlink(document,
                        "https://www.gemboxsoftware.com/document/free-version",
                        "here"),
                    New Run(document, "."))),
            New Run(document, " on the same machine.")))

        document.Sections.Add(
            New Section(document,
                New Paragraph(document,
                    New Run(document, "GemBox.Document Free"),
                    New Note(document, NoteType.Endnote,
                        "Limited to 20 paragraphs."),
                    New Run(document, " is free of charge while GemBox.Document"),
                    New Run(document, " Professional is a commercial version"),
                    New Run(document, " licensed per developer."),
                    New Run(document, " Server deployment is royalty free."))))

        document.Save("Footnotes and Endnotes.docx")

    End Sub

End Module
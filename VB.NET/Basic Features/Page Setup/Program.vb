Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "First section's page.")))

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "Second section's page, first paragraph."),
                New Paragraph(document, "Second section's page, second paragraph."),
                New Paragraph(document, "Second section's page, third paragraph.")))

        Dim pageSetup1 = document.Sections(0).PageSetup
        Dim pageSetup2 = document.Sections(1).PageSetup

        ' Set page orientation.
        pageSetup1.Orientation = Orientation.Landscape

        ' Set page margins.
        pageSetup1.PageMargins.Top = 10
        pageSetup1.PageMargins.Bottom = 10

        ' Set paper type.
        pageSetup2.PaperType = PaperType.A5

        ' Set line numbering.
        pageSetup2.LineNumberRestartSetting = LineNumberRestartSetting.NewPage

        document.Save("Page Setup.docx")

    End Sub
End Module
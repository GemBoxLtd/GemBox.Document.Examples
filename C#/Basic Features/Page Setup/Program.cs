using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "First section's page.")));

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "Second section's page, first paragraph."),
                new Paragraph(document, "Second section's page, second paragraph."),
                new Paragraph(document, "Second section's page, third paragraph.")));

        var pageSetup1 = document.Sections[0].PageSetup;
        var pageSetup2 = document.Sections[1].PageSetup;

        // Set page orientation.
        pageSetup1.Orientation = Orientation.Landscape;

        // Set page margins.
        pageSetup1.PageMargins.Top = 10;
        pageSetup1.PageMargins.Bottom = 10;

        // Set paper type.
        pageSetup2.PaperType = PaperType.A5;

        // Set line numbering.
        pageSetup2.LineNumberRestartSetting = LineNumberRestartSetting.NewPage;

        document.Save("Page Setup.docx");
    }
}

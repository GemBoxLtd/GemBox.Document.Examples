using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        // First section
        var section1 = new Section(document);
        document.Sections.Add(section1);

        var header1 = new HeaderFooter(document, HeaderFooterType.HeaderDefault);
        section1.HeadersFooters.Add(header1);

        var pictureWatermark = new PictureWatermark(document, new Picture(document, "Acme.jpg"));
        header1.Watermark = pictureWatermark; // Assign watermark to the header.
        pictureWatermark.AutoScale(); // Scale the picture to fit the page.
        pictureWatermark.Washout = true;

        // Second section
        var section2 = new Section(document);
        document.Sections.Add(section2);

        var header2 = new HeaderFooter(document, HeaderFooterType.HeaderDefault);
        section2.HeadersFooters.Add(header2);

        var textWatermark = new TextWatermark(document, "Acme corporation");
        header2.Watermark = textWatermark;
        textWatermark.SetDiagonal();
        textWatermark.Color = Color.Red;
        textWatermark.Semitransparent = true;

        document.Save("Watermarks.docx");
    }
}
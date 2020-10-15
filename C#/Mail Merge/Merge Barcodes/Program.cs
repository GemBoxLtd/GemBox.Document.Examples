using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Create a new document.
        var document = new DocumentModel();
        document.DefaultParagraphFormat.Alignment = HorizontalAlignment.Center;

        // Create a barcode merge field that will display the value in a barcode font.
        var barcodeField = new Field(document, FieldType.MergeField, "Barcode", "«Barcode»")
        {
            CharacterFormat =
            {
                FontName = "Code 128",
                Size = 80
            }
        };

        // Create a label merge field that will display the value with a '*' character as the prefix and suffix.
        var labelField = new Field(document, FieldType.MergeField, @"Label \b * \f *", "*«Label»*")
        {
            CharacterFormat =
            {
                FontName = "Arial Black",
                Size = 20,
                FontColor = Color.Red
            }
        };

        // Add merge fields to the document.
        document.Sections.Add(
            new Section(document,
                new Paragraph(document, barcodeField),
                new Paragraph(document, labelField)));

        document.MailMerge.Execute(new { Barcode = "1234567890", Label = "1234567890" });

        document.Save("Barcode Output.docx");
    }
}
using System.Text;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        var document = new DocumentModel();

        var qrCodeValue = "1234567890";
        var qrCodeField = new Field(document, FieldType.DisplayBarcode, $"{qrCodeValue} QR");

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, qrCodeField)));

        document.Save("QR Code Output.pdf");
    }

    static void Example2()
    {
        var document = new DocumentModel();

        var ean13 = CreateBarcodeField(
            document,
            barcodeType: "EAN13",
            barcodeValue: "5901234123457",
            heightInPoints: 100,
            showLabel: true);

        var upca = CreateBarcodeField(
            document,
            barcodeType: "UPCA",
            barcodeValue: "123456789104",
            showLabel: true);

        var code128 = CreateBarcodeField(
            document,
            barcodeType: "Code128",
            barcodeValue: "012345678",
            foregroundColor:"0x2572FF",
            backgroundColor:"0xffb225");


        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "EAN13 Code:"),
                new Paragraph(document, ean13),
                new Paragraph(document, "UPCA Code:"),
                new Paragraph(document, upca),
                new Paragraph(document, "Code 128:"),
                new Paragraph(document, code128)));

        document.Save("Barcodes.pdf");
    }

    static Field CreateBarcodeField(DocumentModel document, string barcodeType, string barcodeValue,
        int? heightInPoints = null, string foregroundColor = null,
        string backgroundColor = null, bool showLabel = false)
    {
        var instructionText = new StringBuilder();
        instructionText.Append(barcodeValue).Append(' ').Append(barcodeType);

        if (heightInPoints.HasValue)
            instructionText.Append(" \\h ").Append(LengthUnitConverter.Convert(heightInPoints.Value, LengthUnit.Point, LengthUnit.Twip));
        if (foregroundColor != null)
            instructionText.Append(" \\f ").Append(foregroundColor);
        if (backgroundColor != null)
            instructionText.Append(" \\b ").Append(backgroundColor);
        if (showLabel)
            instructionText.Append(" \\t");

        return new Field(document, FieldType.DisplayBarcode, instructionText.ToString());
    }
}
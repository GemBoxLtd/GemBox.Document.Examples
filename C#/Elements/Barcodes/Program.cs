using GemBox.Document;
using System.Collections.Generic;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

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
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var barcodeType = "CODE39";

        var barcodeValues = new Dictionary<string, string>()
        {
            ["CODE39"] = "GEMBOX123",
            ["CODE128"] = "GemBox-Document-123",
            ["EAN8"] = "9638507",
            ["EAN13"] = "490123456789",
            ["JAN8"] = "9638507",
            ["JAN13"] = "490123456789",
            ["UPCA"] = "036000291452",
            ["NW7"] = "123456"
        };
        var barcodeValue = barcodeValues[barcodeType];

        var barcodeField = new Field(document, FieldType.DisplayBarcode, $"{barcodeValue} {barcodeType}");

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, $"Barcode '{barcodeType}' with value '{barcodeValue}':"),
                new Paragraph(document, barcodeField)));

        document.Save("Barcode Output.docx");
    }

    static void Example3()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var foregroundSwitch = @"\f 800000";
        var backgroundSwitch = @"\b D3D3D3";
        var showTextSwitch = @"\t";
        var heightSwitch = @"\h 3000";

        var barcodeSwitches = $" {foregroundSwitch} {backgroundSwitch} {showTextSwitch} {heightSwitch}";

        var ean13 = new Field(document, FieldType.DisplayBarcode, "5901234123457 EAN13" + barcodeSwitches);
        var upca = new Field(document, FieldType.DisplayBarcode, "123456789104 UPCA" + barcodeSwitches);
        var code128 = new Field(document, FieldType.DisplayBarcode, "012345678 CODE128" + barcodeSwitches);

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "EAN13:"),
                new Paragraph(document, ean13),
                new Paragraph(document, "UPCA:"),
                new Paragraph(document, upca),
                new Paragraph(document, "CODE128:"),
                new Paragraph(document, code128)));

        document.Save("Formatted Barcodes.pdf");
    }
}

using System.Linq;
using GemBox.Document;
using GemBox.Document.Drawing;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("DigitalSignature.docx");

        // Get placeholder where signature should be visualized.
        // Signature line was added with: Microsoft Word => Insert => Signature Line
        // By default it'll have "Microsoft Office Signature Line..." description.
        var signatureLine = document.GetChildElements(true).OfType<DrawingElement>().FirstOrDefault(
            de => de.Metadata.Description == "Microsoft Office Signature Line...");

        // Create visual representation of digital signature.
        var signature = new Picture(document, "GemBoxSignature.png");

        // Position signature image in a signature line.
        // Image will be placed 1.5cm right and 0.5cm below the top-left corner of signature line.
        signature.Layout = Layout.Floating(
            new HorizontalPosition(1.5, LengthUnit.Centimeter, HorizontalPositionAnchor.Page),
            new VerticalPosition(0.5, LengthUnit.Centimeter, VerticalPositionAnchor.Page),
            signature.Layout.Size);

        var options = new PdfSaveOptions()
        {
            DigitalSignature =
            {
                CertificatePath = "GemBoxExampleExplorer.pfx",
                CertificatePassword = "GemBoxPassword",
                SignatureLine = signatureLine,
                Signature = signature
            }
        };

        document.Save("PDF Digital Signature.pdf", options);
    }
}
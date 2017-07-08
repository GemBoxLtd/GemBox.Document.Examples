using System;
using System.IO;
using System.Linq;
using GemBox.Document;
using GemBox.Document.Drawing;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("DigitalSignature.docx");

        string pathToResources = "Resources";

        // Signature line added with MS Word -> Insert tab -> Signature Line button by default has description 'Microsoft Office Signature Line...'.
        var signatureLine = document.GetChildElements(true).OfType<DrawingElement>().FirstOrDefault(
            de => de.Metadata.Description == "Microsoft Office Signature Line...");

        var signature = new Picture(document, Path.Combine(pathToResources, "GemBoxSignature.png"));

        // Signature in this document will be 1.5 cm right of TopLeft position of signature line 
        // and 0.5 cm below of TopLeft position of signature line.
        signature.Layout = Layout.Floating(
            new HorizontalPosition(1.5, LengthUnit.Centimeter, HorizontalPositionAnchor.Page),
            new VerticalPosition(0.5, LengthUnit.Centimeter, VerticalPositionAnchor.Page),
            signature.Layout.Size);

        var options = new PdfSaveOptions()
        {
            DigitalSignature =
            {
                CertificatePath = Path.Combine(pathToResources, "GemBoxSampleExplorer.pfx"),
                CertificatePassword = "GemBoxPassword",
                // Placeholder where signature should be visualized.
                SignatureLine = signatureLine,
                // Visual representation of digital signature.
                Signature = signature
            }
        };

        document.Save("PDF Digital Signature.pdf", options);
    }
}

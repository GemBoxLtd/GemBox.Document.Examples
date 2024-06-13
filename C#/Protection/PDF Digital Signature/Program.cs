using GemBox.Document;
using GemBox.Pdf.Forms;
using GemBox.Pdf.Security;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");

        // Create visual representation of digital signature.
        var signature = new Picture(document, "GemBoxSignature.png");

        // Position signature image at the end of the document.
        var lastSection = document.Sections[document.Sections.Count - 1];
        lastSection.Blocks.Add(new Paragraph(document, signature));

        var options = new PdfSaveOptions()
        {
            DigitalSignature =
            {
                CertificatePath = "GemBoxECDsa521.pfx",
                CertificatePassword = "GemBoxPassword",
                Signature = signature,
                IsAdvancedElectronicSignature = true
            }
        };

        document.Save("PDF Digital Signature.pdf", options);
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");

        // Create visual representation of digital signature.
        var signature = new Picture(document, "GemBoxSignature.png");

        // Position signature image at the end of the document.
        var lastSection = document.Sections[document.Sections.Count - 1];
        lastSection.Blocks.Add(new Paragraph(document, signature));

        // If using the Professional version, put your serial key below.
        GemBox.Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Get a digital ID from PKCS#12/PFX file.
        var digitalId = new PdfDigitalId("GemBoxECDsa521.pfx", "GemBoxPassword");

        // Create a PDF signer that will create PAdES B-LTA level signature.
        var signer = new PdfSigner(digitalId);

        // PdfSigner should create CAdES-equivalent signature.
        signer.SignatureFormat = PdfSignatureFormat.CAdES;

        // PdfSigner will embed a timestamp created by freeTSA.org Time Stamp Authority in the signature.
        signer.Timestamper = new PdfTimestamper("https://freetsa.org/tsr");

        // Make sure that all properties specified on PdfSigner are according to PAdES B-LTA level.
        signer.SignatureLevel = PdfSignatureLevel.PAdES_B_LTA;

        // Inject PdfSigner from GemBox.Pdf into
        // PdfDigitalSignatureSaveOptions from GemBox.Document.
        var signatureOptions = PdfDigitalSignatureSaveOptions.FromSigner(
            () => signer.SignatureFormat.ToString(),
            () => signer.EstimatedSignatureContentsLength,
            signer.ComputeSignature);

        signatureOptions.Signature = signature;

        var options = new PdfSaveOptions()
        {
            DigitalSignature = signatureOptions
        };

        document.Save("PAdES B-LTA.pdf", options);

        using (var pdfDocument = GemBox.Pdf.PdfDocument.Load("PAdES B-LTA.pdf"))
        {
            var signatureField = (PdfSignatureField)pdfDocument.Form.Fields[0];

            // Download validation-related information for the signature and the signature's timestamp and embed it in the PDF file.
            // This will make the signature "LTV enabled".
            pdfDocument.SecurityStore.AddValidationInfo(signatureField.Value);

            // Add an invisible signature field to the PDF document that will hold the document timestamp.
            var timestampField = pdfDocument.Form.Fields.AddSignature();

            // Initiate timestamping of a PDF file with the specified timestamper.
            timestampField.Timestamp(signer.Timestamper);

            // Save any changes done to the PDF file that were done since the last time Save was called and
            // finish timestamping of a PDF file.
            pdfDocument.Save();
        }
    }
}

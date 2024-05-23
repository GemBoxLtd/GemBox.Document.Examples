using GemBox.Document;
using GemBox.Document.Security;

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

        var saveOptions = new DocxSaveOptions();
        saveOptions.DigitalSignatures.Add(new DocxDigitalSignatureSaveOptions()
        {
            CertificatePath = "GemBoxECDsa521.pfx",
            CertificatePassword = "GemBoxPassword"
        });

        document.Save("DOCX Digital Signature.docx", saveOptions);
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");

        var signature1 = new DocxDigitalSignatureSaveOptions()
        {
            CertificatePath = "GemBoxECDsa521.pfx",
            CertificatePassword = "GemBoxPassword",
            CommitmentType = DigitalSignatureCommitmentType.Created,
            SignerRole = "Developer"
        };
        // Embed intermediate certificate.
        signature1.Certificates.Add(new Certificate("GemBoxECDsa.crt"));

        var signature2 = new DocxDigitalSignatureSaveOptions()
        {
            CertificatePath = "GemBoxRSA4096.pfx",
            CertificatePassword = "GemBoxPassword",
            CommitmentType = DigitalSignatureCommitmentType.Approved,
            SignerRole = "Manager"
        };
        // Embed intermediate certificate.
        signature2.Certificates.Add(new Certificate("GemBoxRSA.crt"));

        var saveOptions = new DocxSaveOptions();
        saveOptions.DigitalSignatures.Add(signature1);
        saveOptions.DigitalSignatures.Add(signature2);

        document.Save("DOCX Digital Signatures.docx", saveOptions);
    }
}

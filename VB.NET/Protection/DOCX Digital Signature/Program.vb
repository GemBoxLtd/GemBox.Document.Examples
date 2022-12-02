Imports GemBox.Document
Imports GemBox.Document.Security

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim document = DocumentModel.Load("Reading.docx")

        Dim saveOptions As New DocxSaveOptions()
        saveOptions.DigitalSignatures.Add(New DocxDigitalSignatureSaveOptions() With
        {
            .CertificatePath = "GemBoxECDsa521.pfx",
            .CertificatePassword = "GemBoxPassword"
        })

        document.Save("DOCX Digital Signature.docx", saveOptions)
    End Sub

    Sub Example2()
        Dim document = DocumentModel.Load("Reading.docx")

        Dim signature1 As New DocxDigitalSignatureSaveOptions() With
        {
            .CertificatePath = "GemBoxECDsa521.pfx",
            .CertificatePassword = "GemBoxPassword",
            .CommitmentType = DigitalSignatureCommitmentType.Created,
            .SignerRole = "Developer"
        }
        ' Embed intermediate certificate.
        signature1.Certificates.Add(New Certificate("GemBoxECDsa.crt"))

        Dim signature2 As New DocxDigitalSignatureSaveOptions() With
        {
            .CertificatePath = "GemBoxRSA4096.pfx",
            .CertificatePassword = "GemBoxPassword",
            .CommitmentType = DigitalSignatureCommitmentType.Approved,
            .SignerRole = "Manager"
        }
        ' Embed intermediate certificate.
        signature2.Certificates.Add(New Certificate("GemBoxRSA.crt"))

        Dim saveOptions As New DocxSaveOptions()
        saveOptions.DigitalSignatures.Add(signature1)
        saveOptions.DigitalSignatures.Add(signature2)

        document.Save("DOCX Digital Signatures.docx", saveOptions)
    End Sub
End Module
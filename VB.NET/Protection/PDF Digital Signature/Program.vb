Imports System.Linq
Imports GemBox.Document
Imports GemBox.Document.Drawing

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("DigitalSignature.docx")

        ' Get placeholder where signature should be visualized.
        ' Signature line was added with: Microsoft Word => Insert => Signature Line
        ' By default it'll have "Microsoft Office Signature Line..." description.
        Dim signatureLine As DrawingElement = document.GetChildElements(True).OfType(Of DrawingElement)().FirstOrDefault(
            Function(de) de.Metadata.Description = "Microsoft Office Signature Line...")

        ' Create visual representation of digital signature.
        Dim signature As New Picture(document, "GemBoxSignature.png")

        ' Position signature image in a signature line.
        ' Image will be placed 1.5cm right and 0.5cm below the top-left corner of signature line.
        signature.Layout = Layout.Floating(
            New HorizontalPosition(1.5, LengthUnit.Centimeter, HorizontalPositionAnchor.Page),
            New VerticalPosition(0.5, LengthUnit.Centimeter, VerticalPositionAnchor.Page),
            signature.Layout.Size)

        Dim options As New PdfSaveOptions() With
        {
            .DigitalSignature = New PdfDigitalSignatureSaveOptions() With
            {
                .CertificatePath = "GemBoxExampleExplorer.pfx",
                .CertificatePassword = "GemBoxPassword",
                .signatureLine = signatureLine,
                .signature = signature
            }
        }

        document.Save("PDF Digital Signature.pdf", options)

    End Sub
End Module
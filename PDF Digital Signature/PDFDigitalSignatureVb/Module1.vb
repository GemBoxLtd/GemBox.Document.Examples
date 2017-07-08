Imports System
Imports System.IO
Imports System.Linq
Imports GemBox.Document
Imports GemBox.Document.Drawing

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("DigitalSignature.docx")

        Dim pathToResources As String = "Resources"

        ' Signature line added with MS Word -> Insert tab -> Signature Line button by default has description 'Microsoft Office Signature Line...'.
        Dim signatureLine As DrawingElement = document.GetChildElements(True).OfType(Of DrawingElement)().FirstOrDefault(
            Function(de) de.Metadata.Description = "Microsoft Office Signature Line...")

        Dim signature = New Picture(document, Path.Combine(pathToResources, "GemBoxSignature.png"))

        ' Signature in this document will be 1.5 cm right of TopLeft position of signature line 
        ' and 0.5 cm below of TopLeft position of signature line.
        signature.Layout = Layout.Floating(
            New HorizontalPosition(1.5, LengthUnit.Centimeter, HorizontalPositionAnchor.Page),
            New VerticalPosition(0.5, LengthUnit.Centimeter, VerticalPositionAnchor.Page),
            signature.Layout.Size)

        Dim options = New PdfSaveOptions()
        Dim digitalSignature = options.DigitalSignature

        digitalSignature.CertificatePath = Path.Combine(pathToResources, "GemBoxSampleExplorer.pfx")
        digitalSignature.CertificatePassword = "GemBoxPassword"
        ' Placeholder where signature should be visualized.
        digitalSignature.SignatureLine = signatureLine
        ' Visual representation of digital signature.
        digitalSignature.Signature = signature

        document.Save("PDF Digital Signature.pdf", options)

    End Sub

End Module
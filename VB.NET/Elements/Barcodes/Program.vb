Imports GemBox.Document
Imports System.Collections.Generic

Module Program

    Sub Main()
        Example1()
        Example2()
        Example3()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim qrCodeValue = "1234567890"
        Dim qrCodeField As New Field(document, FieldType.DisplayBarcode, $"{qrCodeValue} QR")

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, qrCodeField)))

        document.Save("QR Code Output.pdf")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim barcodeType = "CODE39"

        Dim barcodeValues As New Dictionary(Of String, String)() From
        {
            {"CODE39", "GEMBOX123"},
            {"CODE128", "GemBox-Document-123"},
            {"EAN8", "9638507"},
            {"EAN13", "490123456789"},
            {"JAN8", "9638507"},
            {"JAN13", "490123456789"},
            {"UPCA", "036000291452"},
            {"NW7", "123456"}
        }
        Dim barcodeValue = barcodeValues(barcodeType)

        Dim barcodeField As New Field(document, FieldType.DisplayBarcode, $"{barcodeValue} {barcodeType}")

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, $"Barcode '{barcodeType}' with value '{barcodeValue}':"),
                New Paragraph(document, barcodeField)))

        document.Save("Barcode Output.docx")
    End Sub

    Sub Example3()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim foregroundSwitch = "\f 800000"
        Dim backgroundSwitch = "\b D3D3D3"
        Dim showTextSwitch = "\t"
        Dim heightSwitch = "\h 3000"

        Dim barcodeSwitches = $" {foregroundSwitch} {backgroundSwitch} {showTextSwitch} {heightSwitch}"

        Dim ean13 = New Field(document, FieldType.DisplayBarcode, "5901234123457 EAN13" & barcodeSwitches)
        Dim upca = New Field(document, FieldType.DisplayBarcode, "123456789104 UPCA" & barcodeSwitches)
        Dim code128 = New Field(document, FieldType.DisplayBarcode, "012345678 CODE128" & barcodeSwitches)

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "EAN13:"),
                New Paragraph(document, ean13),
                New Paragraph(document, "UPCA:"),
                New Paragraph(document, upca),
                New Paragraph(document, "CODE128:"),
                New Paragraph(document, code128)))

        document.Save("Formatted Barcodes.pdf")
    End Sub

End Module

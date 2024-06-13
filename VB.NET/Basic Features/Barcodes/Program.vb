Imports GemBox.Document
Imports System.Text

Module Program

    Sub Main()
        Example1()
        Example2()
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

        Dim ean13 = CreateBarcodeField(
            document,
            barcodeType:="EAN13",
            barcodeValue:="5901234123457",
            heightInPoints:=100,
            showLabel:=True)

        Dim upca = CreateBarcodeField(
            document,
            barcodeType:="UPCA",
            barcodeValue:="123456789104",
            showLabel:=True)

        Dim code128 = CreateBarcodeField(
            document,
            barcodeType:="Code128",
            barcodeValue:="012345678",
            foregroundColor:="0xff7225",
            backgroundColor:="0x25b2ff")

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "EAN13:"),
                New Paragraph(document, ean13),
                New Paragraph(document, "UPCA:"),
                New Paragraph(document, upca),
                New Paragraph(document, "Code 128:"),
                New Paragraph(document, code128)))

        document.Save("Barcodes.pdf")
    End Sub

    Function CreateBarcodeField(document As DocumentModel, barcodeType As String, barcodeValue As String,
        Optional heightInPoints As Integer? = Nothing, Optional foregroundColor As String = Nothing,
        Optional backgroundColor As String = Nothing, Optional showLabel As Boolean = False) As Field

        Dim instructionText As New StringBuilder()
        instructionText.Append(barcodeValue).Append(" "c).Append(barcodeType)

        If heightInPoints.HasValue Then instructionText.Append(" \h ").Append(LengthUnitConverter.Convert(heightInPoints.Value, LengthUnit.Point, LengthUnit.Twip))
        If foregroundColor IsNot Nothing Then instructionText.Append(" \f ").Append(foregroundColor)
        If backgroundColor IsNot Nothing Then instructionText.Append(" \b ").Append(backgroundColor)
        If showLabel Then instructionText.Append(" \t")

        Return New Field(document, FieldType.DisplayBarcode, instructionText.ToString())
    End Function

End Module

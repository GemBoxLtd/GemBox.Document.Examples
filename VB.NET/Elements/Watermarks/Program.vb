Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        ' Add the first section.
        Dim section1 As New Section(document)
        document.Sections.Add(section1)

        Dim header1 As New HeaderFooter(document, HeaderFooterType.HeaderDefault)
        section1.HeadersFooters.Add(header1)

        ' Create a picture watermark and scale it to fit the page.
        Dim pictureWatermark As New PictureWatermark(document, New Picture(document, "Acme.jpg"))
        header1.Watermark = pictureWatermark
        pictureWatermark.AutoScale()
        pictureWatermark.Washout = True

        ' Add the second section.
        Dim section2 As New Section(document)
        document.Sections.Add(section2)

        Dim header2 As New HeaderFooter(document, HeaderFooterType.HeaderDefault)
        section2.HeadersFooters.Add(header2)

        ' Create a text watermark and rotate it diagonally.
        Dim textWatermark As New TextWatermark(document, "Acme corporation")
        header2.Watermark = textWatermark
        textWatermark.SetDiagonal()
        textWatermark.Color = Color.Red
        textWatermark.Semitransparent = True

        document.Save("Watermarks.docx")

    End Sub
End Module

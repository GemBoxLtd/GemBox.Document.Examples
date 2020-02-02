Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = New DocumentModel()

        ' First section
        Dim section1 = New Section(document)
        document.Sections.Add(section1)

        Dim header1 = New HeaderFooter(document, HeaderFooterType.HeaderDefault)
        section1.HeadersFooters.Add(header1)

        Dim pictureWatermark = New PictureWatermark(document, New Picture(document, "Acme.jpg"))
        header1.Watermark = pictureWatermark ' Assign watermark to the header.
        pictureWatermark.AutoScale() ' Scale the picture to fit the page.
        pictureWatermark.Washout = True

        ' Second section
        Dim section2 = New Section(document)
        document.Sections.Add(section2)

        Dim header2 = New HeaderFooter(document, HeaderFooterType.HeaderDefault)
        section2.HeadersFooters.Add(header2)

        Dim textWatermark = New TextWatermark(document, "Acme corporation")
        header2.Watermark = textWatermark
        textWatermark.SetDiagonal()
        textWatermark.Color = Color.Red
        textWatermark.Semitransparent = True

        document.Save("Watermarks.docx")

    End Sub
End Module
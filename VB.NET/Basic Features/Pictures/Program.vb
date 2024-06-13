Imports GemBox.Document
Imports System

Module Program

    Sub Main()
        Example1()
        Example2()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        Dim paragraph As New Paragraph(document)
        section.Blocks.Add(paragraph)

        ' Create and add an inline picture with GIF image.
        Dim picture1 As New Picture(document, "Zahnrad.gif", 61, 53, LengthUnit.Pixel)
        paragraph.Inlines.Add(picture1)

        ' Create and add a floating picture with PNG image.
        Dim picture2 As New Picture(document, "Dices.png")
        Dim layout2 As New FloatingLayout(
            New HorizontalPosition(HorizontalPositionType.Left, HorizontalPositionAnchor.Page),
            New VerticalPosition(2, LengthUnit.Inch, VerticalPositionAnchor.Page),
            picture2.Layout.Size)
        layout2.WrappingStyle = TextWrappingStyle.InFrontOfText

        picture2.Layout = layout2
        paragraph.Inlines.Add(picture2)

        ' Create and add a floating picture with SVG image.
        Dim picture3 As New Picture(document, "Graphics1.svg", 400, 200, LengthUnit.Pixel)
        Dim layout3 As New FloatingLayout(
            New HorizontalPosition(3.5, LengthUnit.Inch, HorizontalPositionAnchor.Page),
            New VerticalPosition(2, LengthUnit.Inch, VerticalPositionAnchor.Page),
            picture3.Layout.Size)
        layout3.WrappingStyle = TextWrappingStyle.BehindText

        picture3.Layout = layout3
        paragraph.Inlines.Add(picture3)

        document.Save("Pictures.docx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        Dim picture As New Picture(document, "Jellyfish.jpg")
        section.Blocks.Add(New Paragraph(document, picture))

        Dim pictureLayout = picture.Layout
        Dim pictureSize = pictureLayout.Size

        Dim pageSetup = section.PageSetup
        Dim pageSize = New Size(
            pageSetup.PageWidth - pageSetup.PageMargins.Left - pageSetup.PageMargins.Right,
            pageSetup.PageHeight - pageSetup.PageMargins.Top - pageSetup.PageMargins.Bottom)

        Dim ratioX As Double = pageSize.Width / pictureSize.Width
        Dim ratioY As Double = pageSize.Height / pictureSize.Height
        Dim ratio As Double = Math.Min(ratioX, ratioY)

        ' Resize picture element's size.
        If ratio < 1 Then
            pictureLayout.Size = New Size(pictureSize.Width * ratio, pictureSize.Height * ratio)
        End If

        document.Save("LargePicture.docx")
    End Sub

End Module

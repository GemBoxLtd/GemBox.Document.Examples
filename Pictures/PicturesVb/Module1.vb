Imports System
Imports System.IO
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        Dim pathToResources As String = "Resources"

        Dim section = New Section(document)
        document.Sections.Add(section)

        Dim paragraph As New Paragraph(document)
        section.Blocks.Add(paragraph)

        Dim picture1 As New Picture(document, Path.Combine(pathToResources, "Zahnrad.gif"), 61, 53, LengthUnit.Pixel)
        paragraph.Inlines.Add(picture1)

        paragraph.Inlines.Add(New Run(document, "GemBox.Document is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way."))

        Dim picture2 As New Picture(document, Path.Combine(pathToResources, "Dices.png"))
        picture2.Layout = New FloatingLayout(
        New HorizontalPosition(HorizontalPositionType.Left, HorizontalPositionAnchor.Page),
        New VerticalPosition(2, LengthUnit.Inch, VerticalPositionAnchor.Page),
        picture2.Layout.Size) With {.WrappingStyle = TextWrappingStyle.InFrontOfText}
        paragraph.Inlines.Add(picture2)

        Dim picture3 As New Picture(document, Path.Combine(pathToResources, "Graphics1.wmf"), 378, 189, LengthUnit.Pixel)
        picture3.Layout = New FloatingLayout(
            New HorizontalPosition(3.5, LengthUnit.Inch, HorizontalPositionAnchor.Page),
            New VerticalPosition(2, LengthUnit.Inch, VerticalPositionAnchor.Page),
            picture3.Layout.Size) With {.WrappingStyle = TextWrappingStyle.BehindText}
        paragraph.Inlines.Add(picture3)

        document.Save("Pictures.docx")

    End Sub

End Module
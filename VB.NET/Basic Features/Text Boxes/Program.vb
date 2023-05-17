Imports System.Linq
Imports GemBox.Document
Imports GemBox.Document.Drawing

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("TextBoxes.docx")
        Dim paragraph = document.GetChildElements(True).OfType(Of Paragraph)().First()

        ' Create and add an inline textbox.
        Dim textbox1 As New TextBox(document,
            Layout.Inline(New Size(10, 5, LengthUnit.Centimeter)),
            ShapeType.Diamond)

        textbox1.Blocks.Add(New Paragraph(document, "An inline TextBox created with GemBox.Document."))
        textbox1.Blocks.Add(New Paragraph(document, "It has blue fill and outline."))

        textbox1.Fill.SetSolid(New Color(222, 235, 247))
        textbox1.Outline.Fill.SetSolid(Color.Blue)

        paragraph.Inlines.Insert(0, textbox1)

        ' Create and add a floating textbox.
        Dim textBox2 As New TextBox(document,
            Layout.Floating(
                New HorizontalPosition(HorizontalPositionType.Right, HorizontalPositionAnchor.Margin),
                New VerticalPosition(VerticalPositionType.Bottom, VerticalPositionAnchor.Margin),
                New Size(6, 3, LengthUnit.Centimeter)),
            New Paragraph(document, "A floating TextBox created with GemBox.Document."),
            New Paragraph(document, "It's rotated and has a default fill and outline."))

        textBox2.Layout.Transform.Rotation = 30

        paragraph.Inlines.Add(textBox2)

        document.Save("Text Boxes.docx")

    End Sub
End Module
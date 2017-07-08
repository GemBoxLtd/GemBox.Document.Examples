Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("TextBoxes.docx")

        Dim textBox = New TextBox(document,
            Layout.Floating(
                New HorizontalPosition(HorizontalPositionType.Right, HorizontalPositionAnchor.Margin),
                New VerticalPosition(VerticalPositionType.Bottom, VerticalPositionAnchor.Margin),
                New Size(4, 3.5, LengthUnit.Centimeter)),
            New Paragraph(document, "Text Box created with GemBox.Document."),
            New Paragraph(document, "It has default fill and outline."))

        document.Sections(0).Blocks.Cast(Of Paragraph)(0).Inlines.Add(textBox)

        document.Save("Text Boxes.docx")

    End Sub

End Module
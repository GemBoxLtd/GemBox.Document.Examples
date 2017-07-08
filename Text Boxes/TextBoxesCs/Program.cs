using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("TextBoxes.docx");

        var textBox = new TextBox(document,
            Layout.Floating(
                new HorizontalPosition(HorizontalPositionType.Right, HorizontalPositionAnchor.Margin),
                new VerticalPosition(VerticalPositionType.Bottom, VerticalPositionAnchor.Margin),
                new Size(4, 3.5, LengthUnit.Centimeter)),
            new Paragraph(document, "Text Box created with GemBox.Document."),
            new Paragraph(document, "It has default fill and outline."));

        document.Sections[0].Blocks.Cast<Paragraph>(0).Inlines.Add(textBox);

        document.Save("Text Boxes.docx");
    }
}

using System.Linq;
using GemBox.Document;
using GemBox.Document.Drawing;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("TextBoxes.docx");
        var paragraph = document.GetChildElements(true).OfType<Paragraph>().First();

        // Create and add an inline textbox.
        TextBox textbox1 = new TextBox(document,
            Layout.Inline(new Size(10, 5, LengthUnit.Centimeter)),
            ShapeType.Diamond);

        textbox1.Blocks.Add(new Paragraph(document, "An inline TextBox created with GemBox.Document."));
        textbox1.Blocks.Add(new Paragraph(document, "It has blue fill and outline."));

        textbox1.Fill.SetSolid(new Color(222, 235, 247));
        textbox1.Outline.Fill.SetSolid(Color.Blue);

        paragraph.Inlines.Insert(0, textbox1);

        // Create and add a floating textbox.
        TextBox textBox2 = new TextBox(document,
            Layout.Floating(
                new HorizontalPosition(HorizontalPositionType.Right, HorizontalPositionAnchor.Margin),
                new VerticalPosition(VerticalPositionType.Bottom, VerticalPositionAnchor.Margin),
                new Size(6, 3, LengthUnit.Centimeter)),
            new Paragraph(document, "A floating TextBox created with GemBox.Document."),
            new Paragraph(document, "It has default fill and outline."));

        paragraph.Inlines.Add(textBox2);

        document.Save("Text Boxes.docx");
    }
}
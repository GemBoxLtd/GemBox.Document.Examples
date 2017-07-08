using System;
using System.IO;
using GemBox.Document;
using GemBox.Document.Drawing;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        
        DocumentModel document = new DocumentModel();

        string pathToResources = "Resources";

        var section = new Section(document);
        document.Sections.Add(section);

        // Set page color to test shapes with empty (transparent) fill.
        section.PageSetup.PageColor = Color.LightGray;

        var paragraph = new Paragraph(document);
        section.Blocks.Add(paragraph);

        // Calculate page content size which will be used to specify size for line shapes.
        var pageContentSize = new Size(
            section.PageSetup.PageWidth - section.PageSetup.PageMargins.Left - section.PageSetup.PageMargins.Right,
            section.PageSetup.PageHeight - section.PageSetup.PageMargins.Top - section.PageSetup.PageMargins.Bottom);

        // Add (custom formatted) horizontal line to the center of the page content area.
        var horizontalLine = new Shape(document, ShapeType.Line, Layout.Floating(
            new HorizontalPosition(HorizontalPositionType.Left, HorizontalPositionAnchor.Margin),
            new VerticalPosition(VerticalPositionType.Center, VerticalPositionAnchor.Margin),
            new Size(pageContentSize.Width, 0)));
        horizontalLine.Outline.Width = 5;
        horizontalLine.Outline.Fill.SetSolid(Color.Red);
        paragraph.Inlines.Add(horizontalLine);

        // Add (custom formatted) vertical line to the center of the page content area.
        var verticalLine = new Shape(document, ShapeType.Line, Layout.Floating(
            new HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Margin),
            new VerticalPosition(VerticalPositionType.Top, VerticalPositionAnchor.Margin),
            new Size(0, pageContentSize.Height)));
        verticalLine.Outline.Width = 10;
        verticalLine.Outline.Fill.SetSolid(Color.DarkGreen);
        paragraph.Inlines.Add(verticalLine);

        // Used to specify size for all other shapes.
        var shapeSize = new Size(5, 3, LengthUnit.Centimeter);

        // Add (custom formatted) rounded rectangle to the top-left corner of the page content area.
        var roundedRectangle = new Shape(document, ShapeType.RoundedRectangle, Layout.Floating(
            new HorizontalPosition(HorizontalPositionType.Left, HorizontalPositionAnchor.Margin),
            new VerticalPosition(VerticalPositionType.Top, VerticalPositionAnchor.Margin),
            shapeSize));
        // Radius of the corners is 35% of the rounded rectangle height (since it is smaller than width).
        roundedRectangle.AdjustValues["adj"] = 35000;
        roundedRectangle.Outline.Width = 5;
        roundedRectangle.Outline.Fill.SetSolid(Color.DarkRed);
        roundedRectangle.Fill.SetSolid(Color.Yellow);
        paragraph.Inlines.Add(roundedRectangle);

        // Add (default formatted) rectangle to the bottom-left corner of the page content area.
        var rectangle = new Shape(document, ShapeType.Rectangle, Layout.Floating(
            new HorizontalPosition(HorizontalPositionType.Left, HorizontalPositionAnchor.Margin),
            new VerticalPosition(VerticalPositionType.Bottom, VerticalPositionAnchor.Margin),
            shapeSize));
        paragraph.Inlines.Add(rectangle);

        // Add (custom formatted) oval (ellipse) to the top-right corner of the page content area.
        var oval = new Shape(document, ShapeType.Oval, Layout.Floating(
            new HorizontalPosition(HorizontalPositionType.Right, HorizontalPositionAnchor.Margin),
            new VerticalPosition(VerticalPositionType.Top, VerticalPositionAnchor.Margin),
            shapeSize));
        oval.Outline.Width = 10;
        oval.Outline.Fill.SetSolid(Color.Purple);
        // Fill is empty (transparent), so page color is visible inside the shape.
        oval.Fill.SetEmpty();
        paragraph.Inlines.Add(oval);

        // Add (default formatted) oval text-box to the bottom-right corner of the page content area.
        var textBox = new TextBox(document, Layout.Floating(
            new HorizontalPosition(HorizontalPositionType.Right, HorizontalPositionAnchor.Margin),
            new VerticalPosition(VerticalPositionType.Bottom, VerticalPositionAnchor.Margin),
            shapeSize), ShapeType.Oval);
        textBox.Blocks.Add(new Paragraph(document, "This text is inside oval text-box."));
        paragraph.Inlines.Add(textBox);

        // Add (custom formatted) oval-clipped picture to the center of the page content area.
        // Picture is shown above line shapes because its wrapping style is set to 'in front of text'.
        var picture = new Picture(document, new MemoryStream(File.ReadAllBytes(Path.Combine(pathToResources, "Dices.png"))), PictureFormat.Png, new FloatingLayout(
            new HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Margin),
            new VerticalPosition(VerticalPositionType.Center, VerticalPositionAnchor.Margin),
            shapeSize) { WrappingStyle = TextWrappingStyle.InFrontOfText }, ShapeType.Oval);
        picture.Outline.Width = 5;
        picture.Outline.Fill.SetSolid(Color.Blue);
        // Fill is visible because picture contains transparent pixels.
        picture.Fill.SetSolid(Color.Orange);
        paragraph.Inlines.Add(picture);

        document.Save("Shapes.docx");
    }
}

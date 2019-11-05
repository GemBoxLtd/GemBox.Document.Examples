using GemBox.Document;
using GemBox.Document.Drawing;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        var pageSetup = section.PageSetup;
        pageSetup.PageColor = Color.LightGray;

        var paragraph = new Paragraph(document);
        section.Blocks.Add(paragraph);

        // Add yellow rounded rectangle with 35% rounded corners.
        var roundedRectangle = new Shape(document, ShapeType.RoundedRectangle, Layout.Inline(
            new Size(4, 2, LengthUnit.Centimeter)));
        roundedRectangle.AdjustValues["adj"] = 35000;
        roundedRectangle.Outline.Width = 10;
        roundedRectangle.Outline.Fill.SetSolid(Color.Yellow);
        roundedRectangle.Fill.SetEmpty();
        paragraph.Inlines.Add(roundedRectangle);

        // Add red horizontal line.
        var width = pageSetup.PageWidth - pageSetup.PageMargins.Left - pageSetup.PageMargins.Right;
        var horizontalLine = new Shape(document, ShapeType.Line, Layout.Floating(
            new HorizontalPosition(HorizontalPositionType.Left, HorizontalPositionAnchor.Margin),
            new VerticalPosition(6, LengthUnit.Centimeter, VerticalPositionAnchor.InsideMargin),
            new Size(width, 0)));
        horizontalLine.Outline.Width = 4;
        horizontalLine.Outline.Fill.SetSolid(Color.Red);
        paragraph.Inlines.Add(horizontalLine);

        // Add green vertical line.
        var height = pageSetup.PageHeight - pageSetup.PageMargins.Top - pageSetup.PageMargins.Bottom;
        var verticalLine = new Shape(document, ShapeType.Line, Layout.Floating(
            new HorizontalPosition(16, LengthUnit.Centimeter, HorizontalPositionAnchor.Page),
            new VerticalPosition(VerticalPositionType.Top, VerticalPositionAnchor.Margin),
            new Size(85, height)));
        verticalLine.Outline.Width = 8;
        verticalLine.Outline.Fill.SetSolid(Color.DarkGreen);
        paragraph.Inlines.Add(verticalLine);

        // Add blue down arrow with 15% rotation.
        var downArrowLayout = Layout.Floating(
            new HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Margin),
            new VerticalPosition(VerticalPositionType.Center, VerticalPositionAnchor.Margin),
            new Size(8, 12, LengthUnit.Centimeter));
        downArrowLayout.Transform.Rotation = 15;

        var downArrow = new Shape(document, ShapeType.DownArrow, downArrowLayout);
        downArrow.Outline.Width = 3;
        downArrow.Outline.Fill.SetSolid(Color.White);
        downArrow.Fill.SetSolid(new Color(91, 155, 213));
        paragraph.Inlines.Add(downArrow);

        document.Save("Shapes.docx");
    }

    static void Example2()
    {
        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        var paragraph = new Paragraph(document);
        section.Blocks.Add(paragraph);

        var shapes = new Group(document, new FloatingLayout(
            new HorizontalPosition(100, LengthUnit.Point, HorizontalPositionAnchor.Page),
            new VerticalPosition(50, LengthUnit.Point, VerticalPositionAnchor.Page),
            new Size(200, 250, LengthUnit.Point)));
        shapes.Layout.Transform.Rotation = 30;

        paragraph.Inlines.Add(shapes);

        // Add rounded rectangle.
        var roundedRectangle = new Shape(document, ShapeType.RoundedRectangle, Layout.Floating(
            new HorizontalPosition(0, LengthUnit.Point, HorizontalPositionAnchor.TopLeftCorner),
            new VerticalPosition(0, LengthUnit.Point, VerticalPositionAnchor.TopLeftCorner),
            new Size(50, 50, LengthUnit.Point)));

        shapes.Add(roundedRectangle);

        // Add down arrow.
        var downArrowLayout = new Shape(document, ShapeType.DownArrow, Layout.Floating(
            new HorizontalPosition(60, LengthUnit.Point, HorizontalPositionAnchor.TopLeftCorner),
            new VerticalPosition(0, LengthUnit.Point, VerticalPositionAnchor.TopLeftCorner),
            new Size(50, 100, LengthUnit.Point)));

        shapes.Add(downArrowLayout);

        // Add picture.
        var picture = new Picture(document, "Dices.png");
        picture.Layout = Layout.Floating(
            new HorizontalPosition(0, LengthUnit.Point, HorizontalPositionAnchor.TopLeftCorner),
            new VerticalPosition(100, LengthUnit.Point, VerticalPositionAnchor.TopLeftCorner),
            new Size(200, 150, LengthUnit.Point));

        shapes.Add(picture);

        document.Save("GroupShapes.docx");
    }
}
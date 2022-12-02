Imports GemBox.Document
Imports GemBox.Document.Drawing

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        Dim pageSetup = section.PageSetup
        pageSetup.PageColor = Color.LightGray

        Dim paragraph As New Paragraph(document)
        section.Blocks.Add(paragraph)

        ' Add yellow rounded rectangle with 35% rounded corners.
        Dim roundedRectangle As New Shape(document, ShapeType.RoundedRectangle, Layout.Inline(
            New Size(4, 2, LengthUnit.Centimeter)))
        roundedRectangle.AdjustValues("adj") = 35000
        roundedRectangle.Outline.Width = 10
        roundedRectangle.Outline.Fill.SetSolid(Color.Yellow)
        roundedRectangle.Fill.SetEmpty()
        paragraph.Inlines.Add(roundedRectangle)

        ' Add red horizontal line.
        Dim width = pageSetup.PageWidth - pageSetup.PageMargins.Left - pageSetup.PageMargins.Right
        Dim horizontalLine As New Shape(document, ShapeType.Line, Layout.Floating(
            New HorizontalPosition(HorizontalPositionType.Left, HorizontalPositionAnchor.Margin),
            New VerticalPosition(6, LengthUnit.Centimeter, VerticalPositionAnchor.InsideMargin),
            New Size(width, 0)))
        horizontalLine.Outline.Width = 4
        horizontalLine.Outline.Fill.SetSolid(Color.Red)
        paragraph.Inlines.Add(horizontalLine)

        ' Add green vertical line.
        Dim height = pageSetup.PageHeight - pageSetup.PageMargins.Top - pageSetup.PageMargins.Bottom
        Dim verticalLine As New Shape(document, ShapeType.Line, Layout.Floating(
            New HorizontalPosition(16, LengthUnit.Centimeter, HorizontalPositionAnchor.Page),
            New VerticalPosition(VerticalPositionType.Top, VerticalPositionAnchor.Margin),
            New Size(85, height)))
        verticalLine.Outline.Width = 8
        verticalLine.Outline.Fill.SetSolid(Color.DarkGreen)
        paragraph.Inlines.Add(verticalLine)

        ' Add blue down arrow with 15% rotation.
        Dim downArrowLayout = Layout.Floating(
            New HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Margin),
            New VerticalPosition(VerticalPositionType.Center, VerticalPositionAnchor.Margin),
            New Size(8, 12, LengthUnit.Centimeter))
        downArrowLayout.Transform.Rotation = 15

        Dim downArrow As New Shape(document, ShapeType.DownArrow, downArrowLayout)
        downArrow.Outline.Width = 3
        downArrow.Outline.Fill.SetSolid(Color.White)
        downArrow.Fill.SetSolid(New Color(91, 155, 213))
        paragraph.Inlines.Add(downArrow)

        document.Save("Shapes.docx")
    End Sub

    Sub Example2()
        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        Dim paragraph As New Paragraph(document)
        section.Blocks.Add(paragraph)

        Dim shapes As New Group(document, New FloatingLayout(
            New HorizontalPosition(100, LengthUnit.Point, HorizontalPositionAnchor.Page),
            New VerticalPosition(50, LengthUnit.Point, VerticalPositionAnchor.Page),
            New Size(200, 250, LengthUnit.Point)))
        shapes.Layout.Transform.Rotation = 30

        paragraph.Inlines.Add(shapes)

        ' Add rounded rectangle.
        Dim roundedRectangle As New Shape(document, ShapeType.RoundedRectangle, Layout.Floating(
            New HorizontalPosition(0, LengthUnit.Point, HorizontalPositionAnchor.TopLeftCorner),
            New VerticalPosition(0, LengthUnit.Point, VerticalPositionAnchor.TopLeftCorner),
            New Size(50, 50, LengthUnit.Point)))

        shapes.Add(roundedRectangle)

        ' Add down arrow.
        Dim downArrowLayout = New Shape(document, ShapeType.DownArrow, Layout.Floating(
            New HorizontalPosition(60, LengthUnit.Point, HorizontalPositionAnchor.TopLeftCorner),
            New VerticalPosition(0, LengthUnit.Point, VerticalPositionAnchor.TopLeftCorner),
            New Size(50, 100, LengthUnit.Point)))

        shapes.Add(downArrowLayout)

        ' Add picture.
        Dim picture As New Picture(document, "Dices.png")
        picture.Layout = Layout.Floating(
            New HorizontalPosition(0, LengthUnit.Point, HorizontalPositionAnchor.TopLeftCorner),
            New VerticalPosition(100, LengthUnit.Point, VerticalPositionAnchor.TopLeftCorner),
            New Size(200, 150, LengthUnit.Point))

        shapes.Add(picture)

        document.Save("GroupShapes.docx")
    End Sub

End Module
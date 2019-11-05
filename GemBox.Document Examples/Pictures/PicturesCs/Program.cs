using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        var paragraph = new Paragraph(document);
        section.Blocks.Add(paragraph);

        // Create and add an inline picture with GIF image.
        Picture picture1 = new Picture(document, "Zahnrad.gif", 61, 53, LengthUnit.Pixel);
        paragraph.Inlines.Add(picture1);

        // Create and add a floating picture with PNG image.
        Picture picture2 = new Picture(document, "Dices.png");
        FloatingLayout layout2 = new FloatingLayout(
            new HorizontalPosition(HorizontalPositionType.Left, HorizontalPositionAnchor.Page),
            new VerticalPosition(2, LengthUnit.Inch, VerticalPositionAnchor.Page),
            picture2.Layout.Size);
        layout2.WrappingStyle = TextWrappingStyle.InFrontOfText;

        picture2.Layout = layout2;
        paragraph.Inlines.Add(picture2);

        // Create and add a floating picture with WMF image.
        Picture picture3 = new Picture(document, "Graphics1.wmf", 378, 189, LengthUnit.Pixel);
        FloatingLayout layout3 = new FloatingLayout(
            new HorizontalPosition(3.5, LengthUnit.Inch, HorizontalPositionAnchor.Page),
            new VerticalPosition(2, LengthUnit.Inch, VerticalPositionAnchor.Page),
            picture3.Layout.Size);
        layout3.WrappingStyle = TextWrappingStyle.BehindText;

        picture3.Layout = layout3;
        paragraph.Inlines.Add(picture3);

        document.Save("Pictures.docx");
    }
}
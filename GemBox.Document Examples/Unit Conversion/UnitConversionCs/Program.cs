using System;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");
        var pageSetup = document.Sections[0].PageSetup;

        double widthInPoints = pageSetup.PageWidth;
        double heightInPoints = pageSetup.PageHeight;

        Console.WriteLine("Document's page size in different units:");

        foreach (LengthUnit unit in Enum.GetValues(typeof(LengthUnit)))
        {
            double convertedWidth = LengthUnitConverter.Convert(widthInPoints, LengthUnit.Point, unit);
            double convertedHeight = LengthUnitConverter.Convert(heightInPoints, LengthUnit.Point, unit);
            Console.WriteLine($"{convertedWidth} x {convertedHeight} {unit.ToString().ToLowerInvariant()}");
        }
    }
}
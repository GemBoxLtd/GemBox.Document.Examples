using System;
using System.Text;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("Reading.docx");

        StringBuilder sb = new StringBuilder();

        sb.AppendLine("Page size (width X height):");

        double width = document.Sections[0].PageSetup.PageWidth;
        double height = document.Sections[0].PageSetup.PageHeight;

        foreach (LengthUnit unit in Enum.GetValues(typeof(LengthUnit)))
        {
            sb.AppendFormat(
                "{0} X {1} {2}",
                LengthUnitConverter.Convert(width, LengthUnit.Point, unit),
                LengthUnitConverter.Convert(height, LengthUnit.Point, unit),
                unit.ToString().ToLowerInvariant());

            sb.AppendLine();
        }

        Console.WriteLine(sb.ToString());
    }
}
